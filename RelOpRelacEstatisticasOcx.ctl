VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRelacEstatisticasOcx 
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   ScaleHeight     =   6660
   ScaleWidth      =   7980
   Begin VB.ComboBox Estatisticas 
      Height          =   315
      ItemData        =   "RelOpRelacEstatisticasOcx.ctx":0000
      Left            =   1560
      List            =   "RelOpRelacEstatisticasOcx.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1020
      Width           =   3015
   End
   Begin VB.Frame FrameTipoRelacionamento 
      Caption         =   "Tipo de Relacionamento"
      Height          =   1095
      Left            =   240
      TabIndex        =   33
      Top             =   5400
      Width           =   5325
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
         TabIndex        =   11
         Top             =   285
         Width           =   1620
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
         TabIndex        =   12
         Top             =   615
         Width           =   1050
      End
      Begin VB.ComboBox TipoRelacionamento 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   2550
      End
   End
   Begin VB.Frame FrameTipoCliente 
      Caption         =   "Tipo de Cliente"
      Height          =   1095
      Left            =   240
      TabIndex        =   32
      Top             =   4200
      Width           =   5325
      Begin VB.ComboBox TipoCliente 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   585
         Width           =   2550
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
         TabIndex        =   9
         Top             =   615
         Width           =   1050
      End
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
         TabIndex        =   8
         Top             =   285
         Width           =   1620
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5640
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRelacEstatisticasOcx.ctx":0070
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRelacEstatisticasOcx.ctx":01CA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRelacEstatisticasOcx.ctx":0354
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRelacEstatisticasOcx.ctx":0886
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   240
      TabIndex        =   25
      Top             =   1530
      Width           =   5325
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   315
         Left            =   1590
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   630
         TabIndex        =   2
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   315
         Left            =   4215
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataDe 
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
         Height          =   240
         Left            =   195
         TabIndex        =   29
         Top             =   315
         Width           =   345
      End
      Begin VB.Label LabelDataAte 
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
         Left            =   2835
         TabIndex        =   28
         Top             =   345
         Width           =   360
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
      Left            =   5835
      Picture         =   "RelOpRelacEstatisticasOcx.ctx":0A04
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   870
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelacEstatisticasOcx.ctx":0B06
      Left            =   1440
      List            =   "RelOpRelacEstatisticasOcx.ctx":0B08
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   435
      Width           =   2730
   End
   Begin VB.Frame FrameAtendentes 
      Caption         =   "Atendentes"
      Height          =   900
      Left            =   240
      TabIndex        =   22
      Top             =   3195
      Width           =   5325
      Begin VB.ComboBox AtendenteAte 
         Height          =   315
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox AtendenteDe 
         Height          =   315
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label LabelAtendenteAte 
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
         Left            =   2835
         TabIndex        =   24
         Top             =   420
         Width           =   360
      End
      Begin VB.Label LabelAtendenteDe 
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
         Left            =   195
         TabIndex        =   23
         Top             =   420
         Width           =   315
      End
   End
   Begin VB.Frame FrameClientes 
      Caption         =   "Clientes"
      Height          =   900
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   5325
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteDe 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   413
         Width           =   315
      End
      Begin VB.Label LabelClienteAte 
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
         Left            =   2835
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   413
         Width           =   360
      End
   End
   Begin VB.Label LabelEstatiscas 
      AutoSize        =   -1  'True
      Caption         =   "Estatísticas:"
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
      Left            =   360
      TabIndex        =   34
      Top             =   1080
      Width           =   1080
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
      Left            =   750
      TabIndex        =   31
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelacEstatisticasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos de browser
Private WithEvents objEventoClienteInicial As AdmEvento
Attribute objEventoClienteInicial.VB_VarHelpID = -1
Private WithEvents objEventoClienteFinal As AdmEvento
Attribute objEventoClienteFinal.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa eventos de browser
    Set objEventoClienteInicial = New AdmEvento
    Set objEventoClienteFinal = New AdmEvento

    'Carrega a combo Tipo Relacionamento
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, TipoRelacionamento)
    If lErro <> SUCESSO Then gError 131459

    'Carrega a combo Tipo Cliente
    Call Carrega_ComboTipoCliente(TipoCliente)

    'Carrega a combo AtendenteDe
    lErro = CF("Carrega_Atendentes", AtendenteDe)
    If lErro <> SUCESSO Then gError 131460

    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", AtendenteAte)
    If lErro <> SUCESSO Then gError 131461
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 131486

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 131459 To 131461, 131486

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172333)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 131462

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche a Combo Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 131463

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 131462
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 131463

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172334)

    End Select

    Exit Function

End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO***
Public Sub AtendenteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteDe_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteDe)
    If lErro <> SUCESSO Then gError 131464

    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 131465

    Exit Sub

Erro_AtendenteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 131464

        Case 131465
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172335)

    End Select

End Sub

Public Sub AtendenteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteAte_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteAte)
    If lErro <> SUCESSO Then gError 131466

    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 131467

    Exit Sub

Erro_AtendenteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 131466

        Case 131467
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172336)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 131468

    End If

    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 131468

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172337)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    'Se está Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 131469

    End If

    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 131469

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172338)

    End Select

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(DataAte.ClipText) > 0 Then

        'Critica o valor do campo data
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 131470

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 131470

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172339)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        'Critica o valor da data
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 131471

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 131471

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172340)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub TipoRelacionamento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TipoRelacionamento_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, TipoRelacionamento, "AVISO_CRIAR_TIPORELACIONAMENTOCLIENTES")
    If lErro <> SUCESSO Then gError 131472

    Exit Sub

Erro_TipoRelacionamento_Validate:

    Cancel = True

    Select Case gErr

        Case 131472

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172341)

    End Select

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteDe_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteInicial.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else

        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objCliente.sNomeReduzido = ClienteInicial.Text

        sOrdenacao = "Nome Reduzido + Código"

    End If

    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInicial, "", sOrdenacao)

    Exit Sub

Erro_LabelClienteDe_Click:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172342)

    End Select

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteAte_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteFinal.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else

        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objCliente.sNomeReduzido = ClienteFinal.Text

        sOrdenacao = "Nome Reduzido + Código"

    End If

    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFinal, "", sOrdenacao)

    Exit Sub

Erro_LabelClienteAte_Click:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172343)

    End Select

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172344)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172345)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172346)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172347)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a data em uma unidade
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 131473

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 131473
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172348)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a data em uma unidade
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 131474

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 131474
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172349)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a data em uma unidade
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 131475

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 131475
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172350)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a data em uma unidade
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 131476

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 131476
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172351)

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
    If lErro <> SUCESSO Then gError 131477
    
    Select Case Codigo_Extrai(Estatisticas.Text)
    
        Case 1
            gobjRelatorio.sNomeTsk = "RLESTAT"

        Case 2
            gobjRelatorio.sNomeTsk = "RLESTCLI"

        Case 3
            gobjRelatorio.sNomeTsk = "RLESTTP"

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 131477

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172352)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

     'Limpa a tela
    Call LimpaRelatorioRelacEstatisticas

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172353)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 131478

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELACESTATISTICA")

    If vbMsgRes = vbYes Then

        'Exclui o elemento do banco de dados
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 131479

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa a tela
        lErro = LimpaRelatorioRelacEstatisticas()
        If lErro <> SUCESSO Then gError 131480

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 131478
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 131479, 131480

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172354)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 131481

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131482

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Grava no banco de dados
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 131483

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 131484

    'Limpa a tela
    lErro = LimpaRelatorioRelacEstatisticas()
    If lErro <> SUCESSO Then gError 131485

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 131481
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 131482 To 131485

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172355)

    End Select

    Exit Sub

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** FUNÇÕES DE APOIO À TELA - INÍCIO ***
Private Function Define_Padrao() As Long
'Preenche as datas e carrega as combos da tela

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    'Preenche os campos da data com o valor da data atual
    DataDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataAte.Text = Format(gdtDataAtual, "dd/mm/yy")

    AtendenteDe.Text = ""
    AtendenteAte.Text = ""

    Estatisticas.ListIndex = 0

    'defina todos os tipos
    TipoRelacTodos.Value = True
    TipoRelacionamento.Enabled = False

    TipoClienteTodos.Value = True
    TipoCliente.Enabled = False

    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172356)

    End Select

    Exit Function

End Function

Private Function LimpaRelatorioRelacEstatisticas()
'Limpa a tela RelOpRelacClientes

Dim lErro As Long

On Error GoTo Erro_LimpaRelatorioRelacEstatisticas

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 131487

    ComboOpcoes.Text = ""

    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 131488

    LimpaRelatorioRelacEstatisticas = SUCESSO

    Exit Function

Erro_LimpaRelatorioRelacEstatisticas:

    LimpaRelatorioRelacEstatisticas = gErr

    Select Case gErr

        Case 131487, 131488

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172357)

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
    If lErro <> SUCESSO Then gError 131489

    'Preenche a ComboBox Tipo com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        objComboBox.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Exit Sub

Erro_Carrega_ComboTipoCliente:

    Select Case gErr

        Case 131489

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172358)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_De As String
Dim sCliente_Ate As String
Dim iIndice As Integer
Dim sAtendente_De As String
Dim sAtendente_Ate As String
Dim sTipoCliente As String
Dim sTipoRelac As String
Dim sEstatistica As String

On Error GoTo Erro_PreencherRelOp

    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sAtendente_De, sAtendente_Ate, sCliente_De, sCliente_Ate, sTipoCliente, sTipoRelac, sEstatistica)
    If lErro <> SUCESSO Then gError 131490

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 131491

    'Inclui o atendente inicial
    lErro = objRelOpcoes.IncluirParametro("NATENDDE", sAtendente_De)
    If lErro <> AD_BOOL_TRUE Then gError 131492

    'Inclui o atendente inicial
    lErro = objRelOpcoes.IncluirParametro("TATENDDE", AtendenteDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131658

    'Inclui o código final
    lErro = objRelOpcoes.IncluirParametro("NATENDATE", sAtendente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 131493
    
    'Inclui o código final
    lErro = objRelOpcoes.IncluirParametro("TATENDATE", AtendenteAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131659
    
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_De)
    If lErro <> AD_BOOL_TRUE Then gError 131494

    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131660

    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 131495

    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131661

    'Inclui a data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 131496

    'Inclui a data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 131497

    'Inclui o tipo
    lErro = objRelOpcoes.IncluirParametro("TTIPORELACIONAMENTO", sTipoRelac)
    If lErro <> AD_BOOL_TRUE Then gError 131498

    lErro = objRelOpcoes.IncluirParametro("TTIPOCLIENTE", sTipoCliente)
    If lErro <> AD_BOOL_TRUE Then gError 131499

    lErro = objRelOpcoes.IncluirParametro("TESTATISTICA", sEstatistica)
    If lErro <> AD_BOOL_TRUE Then gError 131499

    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAtendente_De, sAtendente_Ate, sCliente_De, sCliente_Ate, sTipoCliente, sTipoRelac)
    If lErro <> SUCESSO Then gError 131500

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 131490 To 131500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172359)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sAtendente_De As String, sAtendente_Ate As String, sCliente_De As String, sCliente_Ate As String, sTipoCliente As String, sTipoRelac As String, sEstatistica As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    'Verifica se o atendente inicial foi preenchido
    If AtendenteDe.Text <> "" Then
        sAtendente_De = CStr(LCodigo_Extrai(AtendenteDe.Text))
    Else
        sAtendente_De = ""
    End If

    'Verifica se o atendente final foi preenchido
    If AtendenteAte.Text <> "" Then
        sAtendente_Ate = CStr(LCodigo_Extrai(AtendenteAte.Text))
    Else
        sAtendente_Ate = ""
    End If

    'Verifica se o atendente inicial é menor que o final, se não for --> ERRO
    If sAtendente_De <> "" And sAtendente_Ate <> "" Then
        If CInt(sAtendente_De) > CInt(sAtendente_Ate) Then gError 131501
    End If

    'Verifica se o Cliente inicial foi preenchido
    If ClienteInicial.Text <> "" Then
        sCliente_De = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_De = ""
    End If

    'Verifica se o Cliente Final foi preenchido
    If ClienteFinal.Text <> "" Then
        sCliente_Ate = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_Ate = ""
    End If

    'Verifica se o Cliente Inicial é menor que o final, se não for --> ERRO
    If sCliente_De <> "" And sCliente_Ate <> "" Then

        If CLng(sCliente_De) > CLng(sCliente_Ate) Then gError 131502

    End If

    'data inicial não pode ser maior que a data final --> ERRO
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then

         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 131503

    End If

    'Se a opção para todos os tipos estiver selecionada
    If TipoRelacTodos.Value = True Then
        sTipoRelac = ""
    Else
        If TipoRelacionamento.Text = "" Then gError 131504
        sTipoRelac = TipoRelacionamento.Text
    End If

    'Se a opção para todos os tipos estiver selecionada
    If TipoClienteTodos.Value = True Then
        sTipoCliente = ""
    Else
        If TipoCliente.Text = "" Then gError 131505
        sTipoCliente = TipoCliente.Text
    End If
    
    If Len(Trim(Estatisticas.Text)) <> 0 Then
        sEstatistica = Estatisticas.Text
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 131501
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_INICIAL_MAIOR_FINAL", gErr)
            AtendenteDe.SetFocus

        Case 131502
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus

        Case 131503
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case 131504
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
            TipoRelacionamento.SetFocus
        
        Case 131505
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
            TipoCliente.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172360)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAtendente_De As String, sAtendente_Ate As String, sCliente_De As String, sCliente_Ate As String, sTipoCliente As String, sTipoRelac As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao

    'Verifica se o Cliente Inicial foi preenchido
    If sCliente_De <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvLong(CLng(sCliente_De))

    End If

    'Verifica se o Cliente Final foi preenchido
    If sCliente_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_Ate))

    End If

    'Verifica se o atendente final foi preenchido
    If sAtendente_De <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Atendente >= " & Forprint_ConvInt(CInt(sAtendente_De))

    End If

    'Verifica se o atendente final foi preenchido
    If sAtendente_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Atendente <= " & Forprint_ConvInt(CInt(sAtendente_Ate))

    End If

    'Verifica se a data inicial foi preenchida
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If

    'Verifica se a data final foi preenchida
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172361)

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
    If lErro <> SUCESSO Then gError 131506

    'Preenche Cliente inicial
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 131507

    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)

    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 131508

    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)

    'Preenche o atendente inicial
    lErro = objRelOpcoes.ObterParametro("NATENDDE", sParam)
    If lErro <> SUCESSO Then gError 131509

    AtendenteDe.Text = sParam
    Call AtendenteDe_Validate(bSGECancelDummy)

    'Preenche o atendente final
    lErro = objRelOpcoes.ObterParametro("NATENDATE", sParam)
    If lErro <> SUCESSO Then gError 131510

    AtendenteAte.Text = sParam
    Call AtendenteAte_Validate(bSGECancelDummy)

    'Preenche a data inicial
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 131511

    Call DateParaMasked(DataDe, CDate(sParam))

    'Preenche a data Final
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 131512

    Call DateParaMasked(DataAte, CDate(sParam))
        
    lErro = objRelOpcoes.ObterParametro("TTIPORELACIONAMENTO", sParam)
    If lErro <> SUCESSO Then gError 131513
    
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
    If lErro <> SUCESSO Then gError 131514
    
    'Preenche o tipo
    If sParam = "" Then
        TipoCliente.ListIndex = -1
        TipoCliente.Enabled = False
        TipoClienteTodos.Value = True
    Else
        TipoClienteApenas.Value = True
        TipoCliente.Enabled = True
        Call Combo_Seleciona_ItemData(TipoCliente, Codigo_Extrai(sParam))
    End If
    
    lErro = objRelOpcoes.ObterParametro("TESTATISTICA", sParam)
    If lErro <> SUCESSO Then gError 131591

    Call Combo_Seleciona_ItemData(Estatisticas, Codigo_Extrai(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 131506 To 131514, 131591

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172362)

    End Select

    Exit Function

End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoClienteInicial = Nothing
    Set objEventoClienteFinal = Nothing

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***
Private Sub objEventoClienteInicial_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteInicial_evSelecao

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteInicial.Text = objCliente.sNomeReduzido

    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoClienteInicial_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172363)

    End Select

End Sub

Private Sub objEventoClienteFinal_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteFinal_evSelecao

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteFinal.Text = objCliente.sNomeReduzido

    Call ClienteFinal_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoClienteFinal_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172364)

    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Relacionamentos x Estatísticas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpRelacEstatisticas"

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




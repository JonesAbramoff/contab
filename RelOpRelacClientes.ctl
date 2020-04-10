VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRelacClientesOcx 
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   ScaleHeight     =   7050
   ScaleWidth      =   8130
   Begin VB.Frame FrameOrigem 
      Caption         =   "Origem"
      Height          =   765
      Left            =   240
      TabIndex        =   39
      Top             =   660
      Width           =   5325
      Begin VB.ComboBox Origem 
         Height          =   315
         ItemData        =   "RelOpRelacClientes.ctx":0000
         Left            =   945
         List            =   "RelOpRelacClientes.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   285
         Width           =   2235
      End
      Begin VB.Label LabelOrigem 
         Caption         =   "Origem:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   41
         Top             =   330
         Width           =   645
      End
   End
   Begin VB.Frame FrameCodigo 
      Caption         =   "Código"
      Height          =   780
      Left            =   240
      TabIndex        =   34
      Top             =   1440
      Width           =   5325
      Begin MSMask.MaskEdBox CodigoDe 
         Height          =   300
         Left            =   600
         TabIndex        =   35
         ToolTipText     =   "Informe o código do relacionamento."
         Top             =   285
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodigoAte 
         Height          =   300
         Left            =   3240
         TabIndex        =   37
         ToolTipText     =   "Informe o código do relacionamento."
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodigoAte 
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
         TabIndex        =   38
         Top             =   315
         Width           =   360
      End
      Begin VB.Label LabelCodigoDe 
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.Frame FrameClientes 
      Caption         =   "Clientes"
      Height          =   900
      Left            =   240
      TabIndex        =   29
      Top             =   3000
      Width           =   5325
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   30
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
         TabIndex        =   31
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         TabIndex        =   33
         Top             =   413
         Width           =   360
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
         TabIndex        =   32
         Top             =   413
         Width           =   315
      End
   End
   Begin VB.Frame FrameAtendentes 
      Caption         =   "Atendentes"
      Height          =   900
      Left            =   240
      TabIndex        =   24
      Top             =   3915
      Width           =   5325
      Begin VB.ComboBox AtendenteDe 
         Height          =   315
         Left            =   600
         TabIndex        =   26
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox AtendenteAte 
         Height          =   315
         Left            =   3240
         TabIndex        =   25
         Top             =   360
         Width           =   1815
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
         TabIndex        =   28
         Top             =   420
         Width           =   315
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
         TabIndex        =   27
         Top             =   420
         Width           =   360
      End
   End
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   660
      Left            =   240
      TabIndex        =   20
      Top             =   4815
      Width           =   5325
      Begin VB.OptionButton Status 
         Caption         =   "Pendente"
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
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton Status 
         Caption         =   "Encerrado"
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
         Index           =   1
         Left            =   2055
         TabIndex        =   22
         Top             =   285
         Width           =   1215
      End
      Begin VB.OptionButton Status 
         Caption         =   "Todos"
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
         Index           =   2
         Left            =   3840
         TabIndex        =   21
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame FrameTipo 
      Caption         =   "Tipo"
      Height          =   1095
      Left            =   240
      TabIndex        =   16
      Top             =   5475
      Width           =   5325
      Begin VB.OptionButton TipoTodos 
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
         TabIndex        =   19
         Top             =   285
         Width           =   1620
      End
      Begin VB.OptionButton TipoApenas 
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
         TabIndex        =   18
         Top             =   615
         Width           =   1050
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   585
         Width           =   2550
      End
   End
   Begin VB.CheckBox ExibirDetalhes 
      Caption         =   "Exibir Detalhes"
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
      TabIndex        =   15
      Top             =   6675
      Width           =   2655
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelacClientes.ctx":0030
      Left            =   1680
      List            =   "RelOpRelacClientes.ctx":0032
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   240
      Width           =   2730
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
      Left            =   6075
      Picture         =   "RelOpRelacClientes.ctx":0034
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   750
      Width           =   1815
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   240
      TabIndex        =   5
      Top             =   2250
      Width           =   5325
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   315
         Left            =   1590
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   11
         Top             =   345
         Width           =   360
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
         TabIndex        =   10
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5880
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRelacClientes.ctx":0136
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRelacClientes.ctx":02B4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRelacClientes.ctx":07E6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRelacClientes.ctx":0970
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
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
      Left            =   990
      TabIndex        =   14
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelacClientesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos de browser
Private WithEvents objEventoCodigoDe As AdmEvento
Attribute objEventoCodigoDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoAte As AdmEvento
Attribute objEventoCodigoAte.VB_VarHelpID = -1
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
    Set objEventoCodigoDe = New AdmEvento
    Set objEventoCodigoAte = New AdmEvento
    Set objEventoClienteInicial = New AdmEvento
    Set objEventoClienteFinal = New AdmEvento
    
    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo)
    If lErro <> SUCESSO Then gError 123427
    
    'Carrega a combo AtendenteDe
    lErro = CF("Carrega_Atendentes", AtendenteDe)
    If lErro <> SUCESSO Then gError 123428
    
    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", AtendenteAte)
    If lErro <> SUCESSO Then gError 123428
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 123403

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 123403, 123427, 123428
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172299)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 123404

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche a Combo Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 123405

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 123404
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 123405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172300)

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
    If lErro <> SUCESSO Then gError 123406
    
    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 102985
    
    Exit Sub
    
Erro_AtendenteDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 123406
        
        Case 102985
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172301)

    End Select

End Sub

Public Sub AtendenteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteAte_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteAte)
    If lErro <> SUCESSO Then gError 123407
    
    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 102986
    
    Exit Sub

Erro_AtendenteAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 123407
        
        Case 102986
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172302)

    End Select

End Sub

Private Sub CodigoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRelacClientes As New ClassRelacClientes

On Error GoTo Erro_CodigoDe_Validate

    'Verifica se o código foi preenchido
    If Len(Trim(CodigoDe.Text)) > 0 Then
        
        objRelacClientes.lCodigo = CodigoDe
        
        'Verifica se o Código é existente
        lErro = RelacionamentoClientes_Le2(objRelacClientes)
        If lErro <> SUCESSO And lErro <> 123472 Then gError 123463
        
        'Se o código não existir --> ERRO
        If lErro = 123472 Then gError 123464
        
    End If
    
    Exit Sub

Erro_CodigoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 123463
        
        Case 123464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTOCLIENENTE_NAO_ENCONTRADO2", gErr, objRelacClientes.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172303)

    End Select
    
    Exit Sub

End Sub

Private Sub CodigoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRelacClientes As New ClassRelacClientes

On Error GoTo Erro_CodigoAte_Validate

    'Verifica se o Código foi preenchido
    If Len(Trim(CodigoAte.Text)) > 0 Then
        
        objRelacClientes.lCodigo = CodigoAte
        
        'Verifica se o código é existente
        lErro = RelacionamentoClientes_Le2(objRelacClientes)
        If lErro <> SUCESSO And lErro <> 123472 Then gError 123465
        
        'Se o código não existir --> ERRO
        If lErro = 123472 Then gError 123466
        
    End If
    
    Exit Sub

Erro_CodigoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 123465
        
        Case 123466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTOCLIENENTE_NAO_ENCONTRADO2", gErr, objRelacClientes.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172304)

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
        If lErro <> SUCESSO Then gError 123408

    End If

    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 123408

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172305)

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
        If lErro <> SUCESSO Then gError 123409

    End If

    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 123409

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172306)

    End Select

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate
    
    If Len(DataAte.ClipText) > 0 Then
        
        'Critica o valor do campo data
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 123410

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 123410

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172307)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        'Critica o valor da data
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 123411

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 123411

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172308)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo, "AVISO_CRIAR_TIPORELACIONAMENTOCLIENTES")
    If lErro <> SUCESSO Then gError 123412
    
    Exit Sub

Erro_Tipo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 123412
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172309)

    End Select

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub LabelCodigoDe_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection

    objRelacionamentoCli.lCodigo = StrParaDbl(CodigoDe.Text)
    
    colSelecao.Add giFilialEmpresa
    
    Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, objEventoCodigoDe)
    
    Exit Sub

End Sub

Private Sub LabelCodigoAte_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection

    objRelacionamentoCli.lCodigo = StrParaDbl(CodigoAte.Text)
    
    colSelecao.Add giFilialEmpresa
    
    Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, objEventoCodigoAte)
    
    Exit Sub

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
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172310)
    
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
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172311)
    
    End Select
    
End Sub

Private Sub TipoTodos_Click()

Dim lErro As Long

On Error GoTo Erro_TipoTodos_Click

    'Desabilita o combotipo
    Tipo.ListIndex = -1
    Tipo.Enabled = False
    
    Exit Sub
    
Erro_TipoTodos_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172312)

    End Select

    Exit Sub

End Sub

Private Sub TipoApenas_Click()

Dim lErro As Long

On Error GoTo Erro_TipoApenas_Click

    'Habilita a ComboTipo
    Tipo.Enabled = True
    
    Exit Sub
    
Erro_TipoApenas_Click:
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172313)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a data em uma unidade
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 123413

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 123413
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172314)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a data em uma unidade
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 123414

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 123414
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172315)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a data em uma unidade
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 123415

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 123415
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172316)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a data em uma unidade
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 123416

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 123416
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172317)

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
    If lErro <> SUCESSO Then gError 123417
    
    'Verifica se a CheckBox ExibirDetalhes está selecionada
    If ExibirDetalhes.Value = 1 Then
        gobjRelatorio.sNomeTsk = "RlCliDet"
    Else
        gobjRelatorio.sNomeTsk = "RlCliRes"
    End If
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 123417

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172318)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
     'Limpa a tela
    lErro = LimpaRelatorioRelacClientes
    If lErro <> SUCESSO Then gError 123419
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 123419
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172319)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 123420

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        'Exclui o elemento do banco de dados
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 123421

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa a tela
        lErro = LimpaRelatorioRelacClientes()
        If lErro <> SUCESSO Then gError 123422
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 123420
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 123421, 123422

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172320)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 123423

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 123424

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Grava no banco de dados
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 123425
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 123426
    
    'Limpa a tela
    lErro = LimpaRelatorioRelacClientes()
    If lErro <> SUCESSO Then gError 123422
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 123423
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 123424, 123425, 123426
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172321)

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
Function Define_Padrao() As Long
'Preenche as datas e carrega as combos da tela

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    'Preenche os campos da data com o valor da data atual
    DataDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataAte.Text = Format(gdtDataAtual, "dd/mm/yy")

    'defina todos os tipos
    TipoTodos.Value = True
    Tipo.Enabled = False
    Status(2).Value = True
    
    ExibirDetalhes.Value = 0
    
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172322)

    End Select

    Exit Function

End Function

Private Function LimpaRelatorioRelacClientes()
'Limpa a tela RelOpRelacClientes

Dim lErro As Long

On Error GoTo Erro_LimpaRelatorioRelacClientes

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 123429
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 123430
    
    LimpaRelatorioRelacClientes = SUCESSO
    
    Exit Function
    
Erro_LimpaRelatorioRelacClientes:

    LimpaRelatorioRelacClientes = gErr
    
    Select Case gErr
    
        Case 123429, 123430
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172323)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_De As String
Dim sCliente_Ate As String
Dim iIndice As Integer
Dim sAtendente_De As String
Dim sAtendente_Ate As String
Dim sCodigo_De As String
Dim sCodigo_Ate As String
Dim sTipo As String
Dim sTipoRelac As String
Dim sCheckTipo As String
Dim sStatus As String

On Error GoTo Erro_PreencherRelOp
   
    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sAtendente_De, sAtendente_Ate, sCliente_De, sCliente_Ate, sCodigo_De, sCodigo_Ate, sTipo, sCheckTipo, sStatus, sTipoRelac)
    If lErro <> SUCESSO Then gError 123431
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 123432
        
    'Inclui o tipo de origem
    lErro = objRelOpcoes.IncluirParametro("TORIGEM", LCodigo_Extrai(Origem.Text))
    If lErro <> AD_BOOL_TRUE Then gError 123467

    'Inclui o código inicial
    lErro = objRelOpcoes.IncluirParametro("TCODIGODE", sCodigo_De)
    If lErro <> AD_BOOL_TRUE Then gError 123433

    'Inclui o código final
    lErro = objRelOpcoes.IncluirParametro("TCODIGOATE", sCodigo_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 123434
         
    'Inclui o atendente inicial
    lErro = objRelOpcoes.IncluirParametro("TATENDDE", sAtendente_De)
    If lErro <> AD_BOOL_TRUE Then gError 123435

    'Inclui o código final
    lErro = objRelOpcoes.IncluirParametro("TATENDATE", sAtendente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 123436
         
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_De)
    If lErro <> AD_BOOL_TRUE Then gError 123437
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 123438
    
    'Inclui a data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 123439
    
    'Inclui a data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 123440

    'Inclui o status
    lErro = objRelOpcoes.IncluirParametro("NSTATUS", sStatus)
    If lErro <> AD_BOOL_TRUE Then gError 123441
     
    'Inclui o tipo
    lErro = objRelOpcoes.IncluirParametro("TTIPORELACIONAMENTO", sTipoRelac)
    If lErro <> AD_BOOL_TRUE Then gError 47637
    
    lErro = objRelOpcoes.IncluirParametro("TTIPO", sTipo)
    If lErro <> AD_BOOL_TRUE Then gError 123442

    lErro = objRelOpcoes.IncluirParametro("TUMTIPO", Tipo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 123443

    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 123444
     
    'Inclui o valor da CheckBox ExibirDetalhes
    lErro = objRelOpcoes.IncluirParametro("NDETALHAR", ExibirDetalhes.Value)
    If lErro <> AD_BOOL_TRUE Then gError 123473
     
    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAtendente_De, sAtendente_Ate, sCliente_De, sCliente_Ate, sCodigo_De, sCodigo_Ate, sStatus, sTipo, sCheckTipo)
    If lErro <> SUCESSO Then gError 123445
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 123431 To 123445, 123467, 123473
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172324)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sAtendente_De As String, sAtendente_Ate As String, sCliente_De As String, sCliente_Ate As String, sCodigo_De As String, sCodigo_Ate As String, sTipo As String, sCheckTipo As String, sStatus As String, sTipoRelac As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'Verifica se o campo código inicial foi preenchido
    If CodigoDe.Text <> "" Then
        sCodigo_De = CStr(CodigoDe.Text)
    Else
        sCodigo_De = ""
    End If
    
    'Verifica se o campo código final foi preenchido
    If CodigoAte.Text <> "" Then
        sCodigo_Ate = CStr(CodigoAte.Text)
    Else
        sCodigo_Ate = ""
    End If
    
    'Verifica se a código inicial é menor que a final, se não for --> ERRO
    If sCodigo_De <> "" And sCodigo_Ate <> "" Then
        If CInt(sCodigo_De) > CInt(sCodigo_Ate) Then gError 123446
    End If
    
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
        If CInt(sAtendente_De) > CInt(sAtendente_Ate) Then gError 123447
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
        
        If CLng(sCliente_De) > CLng(sCliente_Ate) Then gError 123448
        
    End If
    
    'data inicial não pode ser maior que a data final --> ERRO
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 123449
    
    End If
    
    'verifica a opção de status selecionada
    For iIndice = 0 To 2
        If Status(iIndice).Value = True Then sStatus = CStr(iIndice)
    Next

    'Se a opção para todos os tipos estiver selecionada
    If TipoTodos.Value = True Then
        sCheckTipo = "Todos"
        sTipo = ""
    
    'Se a opção para apenas um tipo estiver selecionada
    Else
        If Tipo.Text = "" Then gError 123450
        sCheckTipo = "Um"
        sTipo = CStr(Codigo_Extrai(Tipo.Text))
        sTipoRelac = Tipo.Text
    
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 123446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
            CodigoDe.SetFocus
        
        Case 123447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_INICIAL_MAIOR_FINAL", gErr)
            AtendenteDe.SetFocus
        
        Case 123448
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
        
        Case 123449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
        
        Case 123450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
            Tipo.SetFocus
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172325)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAtendente_De As String, sAtendente_Ate As String, sCliente_De As String, sCliente_Ate As String, sCodigo_De As String, sCodigo_Ate As String, sStatus As String, sTipo As String, sCheckTipo As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
      
    'Verifica se a Origem foi Preenchida se foi coloca-a na Expressão
    If LCodigo_Extrai(Origem.Text) <> 0 Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Origem = " & Forprint_ConvInt(CInt(LCodigo_Extrai(Origem.Text)))
        
    End If
      
    'Verifica se o código inicial foi preenchido
    If sCodigo_De <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo >= " & Forprint_ConvInt(CInt(sCodigo_De))
        
    End If
    
    'Verifica se o Código Final foi preenchido
    If sCodigo_Ate <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo <= " & Forprint_ConvInt(CInt(sCodigo_Ate))
        
    End If
    
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
        
   If Trim(sStatus) <> 2 Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Status = " & Forprint_ConvInt(CInt(sStatus))

    End If
    
    'Se a opção para apenas um tipo estiver selecionada
    If sCheckTipo = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Tipo = " & Forprint_ConvInt(CInt(sTipo))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172326)

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
    If lErro <> SUCESSO Then gError 123451
    
    lErro = objRelOpcoes.ObterParametro("TORIGEM", sParam)
    If lErro <> SUCESSO Then gError 123468
    
    Origem.Text = Origem.List(sParam)
    
    'Preenche o código inicial
    lErro = objRelOpcoes.ObterParametro("TCODIGODE", sParam)
    If lErro <> SUCESSO Then gError 123452
    
    CodigoDe.Text = sParam
    
    'Preenche o código final
    lErro = objRelOpcoes.ObterParametro("TCODIGOATE", sParam)
    If lErro <> SUCESSO Then gError 123453
    
    CodigoAte.Text = sParam
    
    'Preenche Cliente inicial
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 123454
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 123455
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    'Preenche o atendente inicial
    lErro = objRelOpcoes.ObterParametro("TATENDDE", sParam)
    If lErro <> SUCESSO Then gError 123456
    
    AtendenteDe.Text = sParam
    Call AtendenteDe_Validate(bSGECancelDummy)
    
    'Preenche o atendente final
    lErro = objRelOpcoes.ObterParametro("TATENDATE", sParam)
    If lErro <> SUCESSO Then gError 123457
    
    AtendenteAte.Text = sParam
    Call AtendenteAte_Validate(bSGECancelDummy)
        
    'Preenche a data inicial
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 123458

    Call DateParaMasked(DataDe, CDate(sParam))

    'Preenche a data Final
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 123459

    Call DateParaMasked(DataAte, CDate(sParam))
                   
    'Preenche o Status
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then gError 123461

    Status(CInt(sParam)) = True
                   
    lErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lErro <> SUCESSO Then gError 47645
                       
    'Preenche o tipo
    If sParam = "Todos" Then
    
        Tipo.ListIndex = -1
        Tipo.Enabled = False
        TipoTodos.Value = True
    
    Else
    
        'Preenche o tipo
        lErro = objRelOpcoes.ObterParametro("TTIPORELACIONAMENTO", sParam)
        If lErro <> SUCESSO Then gError 123462
                            
        TipoApenas.Value = True
        Tipo.Enabled = True
        Tipo.Text = sParam
        Call Combo_Seleciona(Tipo, iTipo)
        
    End If
    
    'Preenche a CheckBox Exibir Detalhes
    lErro = objRelOpcoes.ObterParametro("NDETALHAR", sParam)
    If lErro <> SUCESSO Then gError 123474
    
    ExibirDetalhes.Value = sParam
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 123451 To 123462, 123468, 123474
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172327)

    End Select

    Exit Function

End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***

Public Function RelacionamentoClientes_Le2(objRelacionamentoClientes As ClassRelacClientes)
'Le da tabela de RelacionamentoClientes todas as informações através do código

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim tRelacionamentoClientes As typeRelacionamentoClientes

On Error GoTo Erro_RelacionamentoClientes_Le2

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 123469
    
    'Inicializa as strings para leitura de dados
    tRelacionamentoClientes.sAssunto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
    tRelacionamentoClientes.sAssunto2 = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Lê no BD o relacionamento com código e filialempresa passados como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Origem, Data, Hora, Tipo, Cliente, FilialCliente, Contato, Atendente, RelacionamentoAnt, Assunto1, Assunto2, Status FROM RelacionamentoClientes WHERE Codigo = ?", tRelacionamentoClientes.lCodigo, tRelacionamentoClientes.iFilialEmpresa, tRelacionamentoClientes.iOrigem, tRelacionamentoClientes.dtData, tRelacionamentoClientes.dHora, tRelacionamentoClientes.lTipo, tRelacionamentoClientes.lCliente, tRelacionamentoClientes.iFilialCliente, tRelacionamentoClientes.iContato, tRelacionamentoClientes.iAtendente, tRelacionamentoClientes.lRelacionamentoAnt, tRelacionamentoClientes.sAssunto1, tRelacionamentoClientes.sAssunto2, tRelacionamentoClientes.iStatus, objRelacionamentoClientes.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 123470
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 123471
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 123472
    
    'Transfere os dados lidos para o obj
    With objRelacionamentoClientes
    
        .lCodigo = tRelacionamentoClientes.lCodigo
        .iFilialEmpresa = tRelacionamentoClientes.iFilialEmpresa
        .iOrigem = tRelacionamentoClientes.iOrigem
        .dtData = tRelacionamentoClientes.dtData
        .dtHora = CDate(tRelacionamentoClientes.dHora)
        .lTipo = tRelacionamentoClientes.lTipo
        .lCliente = tRelacionamentoClientes.lCliente
        .iFilialCliente = tRelacionamentoClientes.iFilialCliente
        .iContato = tRelacionamentoClientes.iContato
        .iAtendente = tRelacionamentoClientes.iAtendente
        .lRelacionamentoAnt = tRelacionamentoClientes.lRelacionamentoAnt
        .sAssunto1 = tRelacionamentoClientes.sAssunto1
        .sAssunto2 = tRelacionamentoClientes.sAssunto2
        .iStatus = tRelacionamentoClientes.iStatus
    
    End With

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    RelacionamentoClientes_Le2 = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Le2:

    RelacionamentoClientes_Le2 = gErr

    Select Case gErr

        Case 123469
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 123470, 123471
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
            
        Case 123472 'Registro não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172328)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigoDe = Nothing
    Set objEventoCodigoAte = Nothing
    Set objEventoClienteInicial = Nothing
    Set objEventoClienteFinal = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***

Private Sub objEventoCodigoDe_evSelecao(obj1 As Object)

Dim objRelacionamentoCli As New ClassRelacClientes
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigoDe_evSelecao

    Set objRelacionamentoCli = obj1
    
    'Preenche o Código Inicial com o Código selecionado
    CodigoDe.Text = objRelacionamentoCli.lCodigo
    
    Call CodigoDe_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoCodigoDe_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172329)
    
    End Select

End Sub

Private Sub objEventoCodigoAte_evSelecao(obj1 As Object)

Dim objRelacionamentoCli As New ClassRelacClientes
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigoAte_evSelecao

    Set objRelacionamentoCli = obj1
    
    'Preenche o Código final com o Código selecionado
    CodigoAte.Text = objRelacionamentoCli.lCodigo
    
    Call CodigoAte_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoCodigoAte_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172330)
    
    End Select

End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172331)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172332)
    
    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Relacionamento Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelacClientes"
    
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



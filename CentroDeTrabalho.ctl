VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CentrodeTrabalho 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   5100
      Index           =   2
      Left            =   135
      TabIndex        =   34
      Top             =   705
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Frame FrameDispMaquina 
         Caption         =   "Disponibilidade em Horas"
         Height          =   2145
         Left            =   105
         TabIndex        =   44
         ToolTipText     =   "Abre tela da Programação de Turno"
         Top             =   2955
         Width           =   9060
         Begin VB.CommandButton BotaoParadas 
            Caption         =   "Paradas não Programadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   2365
            TabIndex        =   67
            ToolTipText     =   "Abre tela da Cadastro de Paradas não Programadas"
            Top             =   1620
            Width           =   2100
         End
         Begin VB.CommandButton BotaoProgramarTurno 
            Caption         =   "Programar Alteração no Turno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   4610
            TabIndex        =   21
            ToolTipText     =   "Abre tela da Programação do Turno"
            Top             =   1605
            Width           =   2100
         End
         Begin VB.CommandButton BotaoProgramarDisponibilidade 
            Caption         =   "Programar Alteração na Disponibilidade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   6855
            TabIndex        =   22
            ToolTipText     =   "Abre tela da Programação da Disponibilidade"
            Top             =   1620
            Width           =   2100
         End
         Begin MSMask.MaskEdBox TurnoMaq 
            Height          =   315
            Left            =   375
            TabIndex        =   59
            Top             =   960
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Format          =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispMaqSeg 
            Height          =   315
            Left            =   2175
            TabIndex        =   52
            Top             =   945
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispMaqQui 
            Height          =   315
            Left            =   5415
            TabIndex        =   53
            Top             =   945
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispMaqSex 
            Height          =   315
            Left            =   6510
            TabIndex        =   54
            Top             =   945
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispMaqSab 
            Height          =   315
            Left            =   7605
            TabIndex        =   55
            Top             =   945
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispMaqTer 
            Height          =   315
            Left            =   3255
            TabIndex        =   56
            Top             =   945
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispMaqQua 
            Height          =   315
            Left            =   4335
            TabIndex        =   57
            Top             =   945
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispMaqDom 
            Height          =   315
            Left            =   1110
            TabIndex        =   58
            Top             =   945
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoCalendario 
            Caption         =   "Calendário"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Abre o Relatório Calendário de Máquina"
            Top             =   1620
            Width           =   2100
         End
         Begin MSFlexGridLib.MSFlexGrid GridDisponibilidadeMaquina 
            Height          =   1305
            Left            =   105
            TabIndex        =   17
            Top             =   240
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   2302
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Máquinas"
         Height          =   2940
         Left            =   105
         TabIndex        =   35
         Top             =   -15
         Width           =   9060
         Begin VB.TextBox Quantidade 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6165
            MaxLength       =   3
            TabIndex        =   36
            Top             =   1635
            Width           =   1575
         End
         Begin VB.CommandButton BotaoMaquinas 
            Caption         =   "Máquinas, Habilidades e Processos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "Abre o Browse de Máquinas, Habilidades e Processos"
            Top             =   2415
            Width           =   2100
         End
         Begin VB.TextBox CodigoItem 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   300
            MaxLength       =   20
            TabIndex        =   38
            Top             =   1170
            Width           =   2010
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1275
            TabIndex        =   37
            Top             =   1665
            Width           =   4200
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2160
            Left            =   105
            TabIndex        =   16
            Top             =   225
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   3810
            _Version        =   393216
            Rows            =   8
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   5205
      Index           =   3
      Left            =   135
      TabIndex        =   60
      Top             =   690
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame9 
         Caption         =   "Operadores de Máquinas"
         Height          =   5115
         Left            =   105
         TabIndex        =   61
         Top             =   -15
         Width           =   9060
         Begin VB.TextBox DescricaoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1275
            TabIndex        =   65
            Top             =   1665
            Width           =   4200
         End
         Begin VB.TextBox CodTipoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   300
            MaxLength       =   20
            TabIndex        =   64
            Top             =   1170
            Width           =   990
         End
         Begin VB.CommandButton BotaoOperadores 
            Caption         =   "Tipo de Mão-de-Obra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   105
            TabIndex        =   63
            ToolTipText     =   "Abre o Browse de Tipos de Mão de Obra"
            Top             =   4530
            Width           =   2100
         End
         Begin VB.TextBox QuantidadeMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6165
            MaxLength       =   3
            TabIndex        =   62
            Top             =   1635
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid GridOperadores 
            Height          =   4215
            Left            =   105
            TabIndex        =   66
            Top             =   225
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   7435
            _Version        =   393216
            Rows            =   8
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   5085
      Index           =   1
      Left            =   150
      TabIndex        =   24
      Top             =   705
      Width           =   9240
      Begin VB.Frame Frame7 
         Caption         =   "Turnos"
         Height          =   990
         Left            =   15
         TabIndex        =   41
         Top             =   2325
         Width           =   2610
         Begin MSMask.MaskEdBox QtdeTurnos 
            Height          =   315
            Left            =   1425
            TabIndex        =   5
            Top             =   195
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   2
            Format          =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HorasTurno 
            Height          =   315
            Left            =   1425
            TabIndex        =   6
            Top             =   585
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            Caption         =   "Horas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   810
            TabIndex        =   43
            Top             =   630
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Quantidade:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   330
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Dias Uteis"
         Height          =   2025
         Left            =   2700
         TabIndex        =   40
         Top             =   1290
         Width           =   1815
         Begin VB.ListBox Dias 
            Height          =   1635
            ItemData        =   "CentroDeTrabalho.ctx":0000
            Left            =   150
            List            =   "CentroDeTrabalho.ctx":0019
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Disponibilidade em Horas"
         Height          =   1740
         Left            =   15
         TabIndex        =   39
         Top             =   3345
         Width           =   9105
         Begin MSMask.MaskEdBox DispCTSeg 
            Height          =   315
            Left            =   2190
            TabIndex        =   51
            Top             =   705
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispCTQui 
            Height          =   315
            Left            =   5430
            TabIndex        =   48
            Top             =   705
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispCTSex 
            Height          =   315
            Left            =   6525
            TabIndex        =   49
            Top             =   705
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispCTSab 
            Height          =   315
            Left            =   7620
            TabIndex        =   50
            Top             =   705
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispCTTer 
            Height          =   315
            Left            =   3270
            TabIndex        =   46
            Top             =   705
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispCTQua 
            Height          =   315
            Left            =   4350
            TabIndex        =   47
            Top             =   705
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DispCTDom 
            Height          =   315
            Left            =   1110
            TabIndex        =   45
            Top             =   705
            Width           =   1100
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.0#"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridTurnosDias 
            Height          =   1335
            Left            =   90
            TabIndex        =   8
            Top             =   255
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   2355
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Competências"
         Height          =   3300
         Left            =   4650
         TabIndex        =   28
         Top             =   15
         Width           =   4470
         Begin VB.CommandButton BotaoCompetencia 
            Caption         =   "Competências"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   135
            TabIndex        =   18
            ToolTipText     =   "Abre o Browse de Competências"
            Top             =   2805
            Width           =   1725
         End
         Begin VB.TextBox CodigoCompetencia 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   600
            MaxLength       =   20
            TabIndex        =   30
            Top             =   1050
            Width           =   1290
         End
         Begin VB.TextBox DescricaoComp 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1890
            TabIndex        =   29
            Top             =   1035
            Width           =   2205
         End
         Begin MSFlexGridLib.MSFlexGrid GridCompetencias 
            Height          =   2415
            Left            =   150
            TabIndex        =   9
            Top             =   285
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   4260
            _Version        =   393216
            Rows            =   8
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Carga Esperada"
         DragMode        =   1  'Automatic
         Height          =   1050
         Left            =   15
         TabIndex        =   25
         Top             =   1290
         Width           =   2610
         Begin MSMask.MaskEdBox CargaMin 
            Height          =   315
            Left            =   1425
            TabIndex        =   3
            Top             =   255
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CargaMax 
            Height          =   315
            Left            =   1425
            TabIndex        =   4
            Top             =   645
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCargaMin 
            Caption         =   "Mínima:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   705
            TabIndex        =   27
            Top             =   270
            Width           =   705
         End
         Begin VB.Label LabelCargaMax 
            Caption         =   "Máxima:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   690
            TabIndex        =   26
            Top             =   675
            Width           =   675
         End
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2310
         Picture         =   "CentroDeTrabalho.ctx":0055
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Numeração Automática"
         Top             =   75
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1455
         TabIndex        =   0
         Top             =   75
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   945
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   495
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelDescricao 
         Caption         =   "Descrição:"
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
         Height          =   315
         Left            =   480
         TabIndex        =   33
         Top             =   960
         Width           =   945
      End
      Begin VB.Label LabelCodigo 
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
         Height          =   315
         Left            =   765
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   105
         Width           =   690
      End
      Begin VB.Label LabelNomeReduzido 
         Caption         =   "Nome Reduzido:"
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
         Height          =   315
         Left            =   0
         TabIndex        =   31
         Top             =   540
         Width           =   1410
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   7290
      ScaleHeight     =   480
      ScaleWidth      =   2085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "CentroDeTrabalho.ctx":013F
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CentroDeTrabalho.ctx":02BD
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CentroDeTrabalho.ctx":07EF
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CentroDeTrabalho.ctx":0979
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5595
      Left            =   90
      TabIndex        =   10
      Top             =   345
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Maquinas, Habilidades e Processos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Operadores"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "CentrodeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iTurnoAlterado As Integer

Dim iFrameAtual As Integer

'Grid de Itens
Dim objGridItens As AdmGrid
Dim iGrid_CodigoItem_Col As Integer
Dim iGrid_DescricaoItem_Col As Integer
Dim iGrid_Quantidade_Col As Integer

'Grid de Competências
Dim objGridCompetencias As AdmGrid
Dim iGrid_CodigoCompetencia_Col As Integer
Dim iGrid_DescricaoComp_Col As Integer

'GridTurnosDias
Dim objGridTurnosDias As AdmGrid
Dim iGrid_DispCTDom_Col As Integer
Dim iGrid_DispCTSeg_Col As Integer
Dim iGrid_DispCTTer_Col As Integer
Dim iGrid_DispCTQua_Col As Integer
Dim iGrid_DispCTQui_Col As Integer
Dim iGrid_DispCTSex_Col As Integer
Dim iGrid_DispCTSab_Col As Integer

'GridDisponibilidadeMaquina
Dim objGridDisponibilidadeMaquina As AdmGrid
Dim iGrid_TurnoMaq_Col As Integer
Dim iGrid_DispMaqDom_Col As Integer
Dim iGrid_DispMaqSeg_Col As Integer
Dim iGrid_DispMaqTer_Col As Integer
Dim iGrid_DispMaqQua_Col As Integer
Dim iGrid_DispMaqQui_Col As Integer
Dim iGrid_DispMaqSex_Col As Integer
Dim iGrid_DispMaqSab_Col As Integer

'Grid de Operadores
Dim objGridOperadores As AdmGrid
Dim iGrid_CodigoTipoMO_Col As Integer
Dim iGrid_DescricaoMo_Col As Integer
Dim iGrid_QuantidadeMO_Col As Integer

Dim gcolMaquinas As Collection
Dim iLinhaAntiga As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoMaquina As AdmEvento
Attribute objEventoMaquina.VB_VarHelpID = -1
Private WithEvents objEventoCompetencia As AdmEvento
Attribute objEventoCompetencia.VB_VarHelpID = -1
Private WithEvents objEventoTipoDeMaodeObra As AdmEvento
Attribute objEventoTipoDeMaodeObra.VB_VarHelpID = -1


Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Centros de Trabalho"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CentrodeTrabalho"

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

Private Sub BotaoOperadores_Click()

Dim lErro As Long
Dim objTiposDeMaodeObras As New ClassTiposDeMaodeObra
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoOperadores_Click

    If Me.ActiveControl Is CodTipoMO Then
    
        objTiposDeMaodeObras.iCodigo = StrParaInt(CodTipoMO.Text)
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridOperadores.Row = 0 Then gError 139084

        objTiposDeMaodeObras.iCodigo = StrParaInt(GridOperadores.TextMatrix(GridOperadores.Row, iGrid_CodigoTipoMO_Col))
        
    End If

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objTiposDeMaodeObras, objEventoTipoDeMaodeObra)

    Exit Sub

Erro_BotaoOperadores_Click:

    Select Case gErr
        
        Case 139084
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144332)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProgramarDisponibilidade_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas
Dim objCTMaqProgDisp As ClassCTMaqProgDisp

On Error GoTo Erro_BotaoProgramarDisponibilidade_Click

    'Verifica se existe uma linha do grid selecionada
    If GridItens.Row = 0 Then gError 137263
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)) = 0 Then gError 137264
    
    'Verifica se o Código do CT está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 137265
    
    objCentrodeTrabalho.lCodigo = StrParaLong(Codigo.Text)
    objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
            
    'Lê o CentrodeTrabalho que está sendo Passado
    lErro = CF("CentrodeTrabalho_Le", objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 134449 Then gError 137266
    
    'se CT não está cadastrado -> Erro
    If lErro <> SUCESSO Then gError 137267
            
     'Inicializa o controle para ler a máquina
    CodigoItem.Text = GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)

    Set objMaquinas = New ClassMaquinas

    'Verifica a existencia da máquina e lê seu NumIntDoc
    lErro = CF("TP_Maquina_Le", CodigoItem, objMaquinas)
    If lErro <> SUCESSO Then gError 137268
    
    'Inicializa o obj da tela a ser chamada
    Set objCTMaqProgDisp = New ClassCTMaqProgDisp
    
    objCTMaqProgDisp.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    objCTMaqProgDisp.lNumIntDocMaq = objMaquinas.lNumIntDoc
    
    'Chama a tela ...
    lErro = Chama_Tela("CTMaquinaProgDisp", objCTMaqProgDisp)
    If lErro <> SUCESSO Then gError 137269
    
    Exit Sub
    
Erro_BotaoProgramarDisponibilidade_Click:

    Select Case gErr
    
        Case 137263
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRIDMAQUINA_NAO_SELECIONADA", gErr)
        
        Case 137264
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 137265
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)
        
        Case 137266, 137268, 137269
            'erro tratado nas rotinas chamadas
        
        Case 137267
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_CADASTRADO", gErr, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144333)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoProgramarTurno_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas
Dim objCTMaqProgTurno As ClassCTMaqProgTurno

On Error GoTo Erro_BotaoProgramarTurno_Click

    'Verifica se existe uma linha do grid selecionada
    If GridItens.Row = 0 Then gError 137270
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)) = 0 Then gError 137271
    
    'Verifica se o Código do CT está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 137272
    
    objCentrodeTrabalho.lCodigo = StrParaLong(Codigo.Text)
    objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
            
    'Lê o CentrodeTrabalho que está sendo Passado
    lErro = CF("CentrodeTrabalho_Le", objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 134449 Then gError 137273
    
    'se CT não está cadastrado -> Erro
    If lErro <> SUCESSO Then gError 137274
            
     'Inicializa o controle para ler a máquina
    CodigoItem.Text = GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)

    Set objMaquinas = New ClassMaquinas

    'Verifica a existencia da máquina e lê seu NumIntDoc
    lErro = CF("TP_Maquina_Le", CodigoItem, objMaquinas)
    If lErro <> SUCESSO Then gError 137275
    
    'Inicializa o obj da tela a ser chamada
    Set objCTMaqProgTurno = New ClassCTMaqProgTurno
    
    objCTMaqProgTurno.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    objCTMaqProgTurno.lNumIntDocMaq = objMaquinas.lNumIntDoc
    
    'Chama a tela ...
    lErro = Chama_Tela("CTMaqProgTurno", objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 137276
    
    Exit Sub
    
Erro_BotaoProgramarTurno_Click:

    Select Case gErr
    
        Case 137270
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRIDMAQUINA_NAO_SELECIONADA", gErr)
        
        Case 137271
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 137272
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)
        
        Case 137273, 137275, 137276
            'erro tratado nas rotinas chamadas
        
        Case 137274
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_CADASTRADO", gErr, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144334)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoCalendario_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas
Dim objCTMaquinas As ClassCTMaquinas

On Error GoTo Erro_BotaoCalendario_Click

    'Verifica se existe uma linha do grid selecionada
    If GridItens.Row = 0 Then gError 137277
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)) = 0 Then gError 137278
    
    'Verifica se o Código do CT está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 137279
    
    objCentrodeTrabalho.lCodigo = StrParaLong(Codigo.Text)
    objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
            
    'Lê o CentrodeTrabalho que está sendo Passado
    lErro = CF("CentrodeTrabalho_Le", objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 134449 Then gError 137280
    
    'se CT não está cadastrado -> Erro
    If lErro <> SUCESSO Then gError 137281
            
     'Inicializa o controle para ler a máquina
    CodigoItem.Text = GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)

    Set objMaquinas = New ClassMaquinas

    'Verifica a existencia da máquina e lê seu NumIntDoc
    lErro = CF("TP_Maquina_Le", CodigoItem, objMaquinas)
    If lErro <> SUCESSO Then gError 137282
    
    'Inicializa o obj da tela a ser chamada
    Set objCTMaquinas = New ClassCTMaquinas
    
    objCTMaquinas.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    objCTMaquinas.lNumIntDocMaq = objMaquinas.lNumIntDoc
    
    'Chama a relatório ...
    lErro = CF("RelCalendarioMaquina_Prepara", objCTMaquinas)
    If lErro <> SUCESSO Then gError 137283
    
    Exit Sub
    
Erro_BotaoCalendario_Click:

    Select Case gErr
    
        Case 137277
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRIDMAQUINA_NAO_SELECIONADA", gErr)
        
        Case 137278
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 137279
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)
        
        Case 137280, 137282, 137283
            'erro tratado nas rotinas chamadas
        
        Case 137281
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_CADASTRADO", gErr, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144335)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo numero disponível para um Centro de Trabalho
    lErro = CF("CentroDeTrabalho_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 134339
    
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 134339
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144336)
    
    End Select

    Exit Sub

End Sub

Private Sub Dias_ItemCheck(Item As Integer)

    iAlterado = REGISTRO_ALTERADO
    
    Call Trata_DiasUteisNoGrid(Item)
    
End Sub


Private Sub GridCompetencias_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCompetencias)
        
End Sub

Private Sub GridDisponibilidadeMaquina_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhaAnterior As Integer
Dim iLinhasExistentesAnterior As Integer

On Error GoTo Erro_GridDisponibilidadeMaquina_KeyDown

    'guarda as linhas do grid antes de apagar
    iLinhaAnterior = GridDisponibilidadeMaquina.Row
    iLinhasExistentesAnterior = objGridDisponibilidadeMaquina.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridDisponibilidadeMaquina)
    
    'se apagou a linha realmente ...
    
    If objGridDisponibilidadeMaquina.iLinhasExistentes < iLinhasExistentesAnterior Then
    
        'apaga o objTurno
        gcolMaquinas.Item(GridItens.Row).colTurnos.Remove iLinhaAnterior
            
    End If
    
    Exit Sub
    
Erro_GridDisponibilidadeMaquina_KeyDown:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144337)
    
    End Select

    Exit Sub
        
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridItens_KeyDown

    'Guarda iLinhasExistentes
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    'Verifica se a Tecla apertada foi Del
    If KeyCode = vbKeyDelete Then

        'Guarda o índice da Linha a ser Excluída
        iLinhaAnterior = GridItens.Row
        
    End If

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    
    'Verifica se a Linha foi realmente excluída
    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        'Remove item da coleção
        gcolMaquinas.Remove iLinhaAnterior
        
    End If
    
    lErro = Preenche_GridDispMaquina()
    If lErro <> SUCESSO Then gError 137284
    
    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 137284
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144338)
    
    End Select

    Exit Sub

End Sub

Private Sub GridOperadores_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridOperadores)

End Sub

Private Sub HorasTurno_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQtdeHorasTotal As Double

On Error GoTo Erro_HorasTurno_Validate

    'Verifica se HorasTurno está preenchida
    If Len(Trim(HorasTurno.Text)) <> 0 And iTurnoAlterado = REGISTRO_ALTERADO Then

        'Critica a Quantidade de Horas
        lErro = Valor_Positivo_Critica(Trim(HorasTurno.Text))
        If lErro <> SUCESSO Then gError 137285
        
        If Len(Trim(QtdeTurnos.Text)) <> 0 Then
        
            dQtdeHorasTotal = StrParaInt(QtdeTurnos.Text) * StrParaDbl(HorasTurno.Text)
            
            'se Quantidade de Horas total dos Turnos > 24 horas == Erro
            If dQtdeHorasTotal > HORAS_DO_DIA Then gError 137286
            
            'preenche o grid Turnos Dias com os valores default
            lErro = Preenche_GridTurnosDias_Padrao()
            If lErro <> SUCESSO Then gError 137287
        
        End If
        
        iTurnoAlterado = 0

    End If

    Exit Sub

Erro_HorasTurno_Validate:

    Cancel = True

    Select Case gErr

        Case 137285, 137287
            'erros tratados nas rotinas chamadas
        
        Case 137286
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDEHORASTURNO_EXCEDE_DIA", gErr, dQtdeHorasTotal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144339)

    End Select

    Exit Sub

End Sub

Private Sub HorasTurno_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HorasTurno, iAlterado)
    
End Sub

Private Sub HorasTurno_Change()

    iAlterado = REGISTRO_ALTERADO
    iTurnoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoTipoDeMaodeObra_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer

On Error GoTo Erro_objEventoTipoDeMaodeObra_evSelecao

    Set objTiposDeMaodeObra = obj1

    'Verifica se há algum Tipo MO repetido no grid
    For iLinha = 1 To objGridOperadores.iLinhasExistentes
        
        If iLinha < GridOperadores.Row Then
                                                
            If GridOperadores.TextMatrix(iLinha, iGrid_CodigoTipoMO_Col) = objTiposDeMaodeObra.iCodigo Then
                CodTipoMO.Text = ""
                gError 139085
                
            End If
                
        End If
                       
    Next
    
    CodTipoMO.Text = CStr(objTiposDeMaodeObra.iCodigo)
    
    If Not (Me.ActiveControl Is CodTipoMO) Then
        GridOperadores.TextMatrix(GridOperadores.Row, iGrid_CodigoTipoMO_Col) = CStr(objTiposDeMaodeObra.iCodigo)
        GridOperadores.TextMatrix(GridOperadores.Row, iGrid_DescricaoMo_Col) = objTiposDeMaodeObra.sDescricao
    End If

    'verifica se precisa preencher o grid com uma nova linha
    If GridOperadores.Row - GridOperadores.FixedRows = objGridOperadores.iLinhasExistentes Then
        objGridOperadores.iLinhasExistentes = objGridOperadores.iLinhasExistentes + 1
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoTipoDeMaodeObra_evSelecao:

    Select Case gErr

        Case 139085
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_REPETIDO", gErr, objTiposDeMaodeObra.iCodigo, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144340)

    End Select

    Exit Sub

End Sub

Private Sub QtdeTurnos_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQtdeHorasTotal As Double

On Error GoTo Erro_QtdeTurnos_Validate

    'Verifica se QtdeTurnos está preenchida
    If Len(Trim(QtdeTurnos.Text)) <> 0 And iTurnoAlterado = REGISTRO_ALTERADO Then

        'Critica a Quantidade de Turnos
        lErro = Inteiro_Critica(Trim(QtdeTurnos.Text))
        If lErro <> SUCESSO Then gError 137288
        
        If Len(Trim(HorasTurno.Text)) <> 0 Then
        
            dQtdeHorasTotal = StrParaInt(QtdeTurnos.Text) * StrParaDbl(HorasTurno.Text)
            
            'se Quantidade de Horas total dos Turnos > 24 horas == Erro
            If dQtdeHorasTotal > HORAS_DO_DIA Then gError 137289
            
            'preenche o grid Turnos Dias com os valores default
            lErro = Preenche_GridTurnosDias_Padrao()
            If lErro <> SUCESSO Then gError 137290
        
        End If
        
        iTurnoAlterado = 0
        
    End If

    Exit Sub

Erro_QtdeTurnos_Validate:

    Cancel = True

    Select Case gErr

        Case 137288, 137290
        
        Case 137289
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDEHORASTURNO_EXCEDE_DIA", gErr, dQtdeHorasTotal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144341)

    End Select

    Exit Sub

End Sub

Private Sub QtdeTurnos_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(QtdeTurnos, iAlterado)
    
End Sub

Private Sub QtdeTurnos_Change()

    iAlterado = REGISTRO_ALTERADO
    iTurnoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If

End Sub

Private Sub TurnoMaq_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TurnoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub


Private Sub TurnoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDisponibilidadeMaquina)

End Sub

Private Sub TurnoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDisponibilidadeMaquina.objControle = TurnoMaq
    lErro = Grid_Campo_Libera_Foco(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
        If Me.ActiveControl Is CodigoItem Then Call BotaoMaquinas_Click
        If Me.ActiveControl Is CodigoCompetencia Then Call BotaoCompetencia_Click
        If Me.ActiveControl Is CodTipoMO Then Call BotaoOperadores_Click
    
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
        
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
    
Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    Set objEventoMaquina = Nothing
    Set objEventoCompetencia = Nothing
    Set objEventoTipoDeMaodeObra = Nothing
       
    Set gcolMaquinas = Nothing
    
    Set objGridItens = Nothing
    Set objGridCompetencias = Nothing
    Set objGridTurnosDias = Nothing
    Set objGridDisponibilidadeMaquina = Nothing
    Set objGridOperadores = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144342)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoMaquina = New AdmEvento
    Set objEventoCompetencia = New AdmEvento
    Set objEventoTipoDeMaodeObra = New AdmEvento
            
    iFrameAtual = 1
    
    'Grid Itens
    Set objGridItens = New AdmGrid
    
    'tela em questão
    Set objGridItens.objForm = Me
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 134340
    
    'Grid Competencias
    Set objGridCompetencias = New AdmGrid
    
    'tela em questão
    Set objGridCompetencias.objForm = Me
        
    lErro = Inicializa_GridCompetencias(objGridCompetencias)
    If lErro <> SUCESSO Then gError 134341

    'Grid TurnosDias
    Set objGridTurnosDias = New AdmGrid

    'tela em questão
    Set objGridTurnosDias.objForm = Me

    lErro = Inicializa_GridTurnosDias(objGridTurnosDias)
    If lErro <> SUCESSO Then gError 137291

    'Grid Disponibilidade Maquina
    Set objGridDisponibilidadeMaquina = New AdmGrid

    'tela em questão
    Set objGridDisponibilidadeMaquina.objForm = Me

    lErro = Inicializa_GridDispMaquina(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then gError 137292
    
    Set gcolMaquinas = New Collection
    
    'Grid Operadores
    Set objGridOperadores = New AdmGrid
    
    'tela em questão
    Set objGridOperadores.objForm = Me
    
    lErro = Inicializa_GridOperadores(objGridOperadores)
    If lErro <> SUCESSO Then gError 139075
    
    FrameDispMaquina.Caption = "Disponibilidade em Horas"
    
    Call Marca_Default_Dias
    
    iAlterado = 0
    iTurnoAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 134340, 134341, 137291, 137292, 139075
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144343)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCentrodeTrabalho Is Nothing) Then

        lErro = Traz_CentrodeTrabalho_Tela(objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134957 Then gError 134342

        If lErro <> SUCESSO Then
                
            If objCentrodeTrabalho.lCodigo > 0 Then
                    
                'Coloca o código do CentrodeTrabalho na tela
                Codigo.Text = objCentrodeTrabalho.lCodigo
                        
            ElseIf Len(Trim(objCentrodeTrabalho.sNomeReduzido)) > 0 Then
                    
                'Coloca o NomeReduzido do CentrodeTrabalho na tela
                NomeReduzido.Text = objCentrodeTrabalho.sNomeReduzido
                    
            End If
    
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 134342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144344)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objMaquinas As ClassMaquinas
Dim objCTMaquinas As ClassCTMaquinas
Dim objCompetencias As ClassCompetencias
Dim objCTCompetencias As ClassCTCompetencias
Dim iDiaDaSemana As Integer
Dim objTurnosDias As ClassTurno
Dim objCTMaqCol As New ClassCTMaquinas
Dim objCTOperadores As ClassCTOperadores

On Error GoTo Erro_Move_Tela_Memoria

    objCentrodeTrabalho.lCodigo = StrParaInt(Codigo.Text)
    objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
    objCentrodeTrabalho.sNomeReduzido = NomeReduzido.Text
    objCentrodeTrabalho.sDescricao = Descricao.Text
    objCentrodeTrabalho.dCargaMin = StrParaDbl(Val(CargaMin.Text) / 100)
    objCentrodeTrabalho.dCargaMax = StrParaDbl(Val(CargaMax.Text) / 100)
    
    objCentrodeTrabalho.iTurnos = StrParaInt(QtdeTurnos.Text)
    objCentrodeTrabalho.dHorasTurno = StrParaDbl(HorasTurno.Text)
    
    For iDiaDaSemana = DOMINGO To SABADO
    
        If Dias.Selected(iDiaDaSemana - 1) = True Then
            objCentrodeTrabalho.iDiaisUteis(iDiaDaSemana) = MARCADO
        Else
            objCentrodeTrabalho.iDiaisUteis(iDiaDaSemana) = DESMARCADO
        End If
    
    Next
    
    'Ir preenchendo a colecao no objCentrodeTrabalho com todas as linhas "existentes"
    'do grid TurnosDias
    For iIndice = 1 To objGridTurnosDias.iLinhasExistentes + 1

        Set objTurnosDias = New ClassTurno
        
        objTurnosDias.iTurno = iIndice
        
        For iDiaDaSemana = DOMINGO To SABADO
        
            objTurnosDias.dQtdHoras(iDiaDaSemana) = StrParaDbl(GridTurnosDias.TextMatrix(iIndice, iDiaDaSemana))
    
        Next
        
        objCentrodeTrabalho.colTurnos.Add objTurnosDias
    
    Next
    
    'Ir preenchendo a colecao no objCentrodeTrabalho com todas as linhas "existentes"
    'do grid CTCompetencias
    For iIndice = 1 To objGridCompetencias.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridCompetencias.TextMatrix(iIndice, iGrid_CodigoCompetencia_Col))) = 0 Then Exit For
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = GridCompetencias.TextMatrix(iIndice, iGrid_CodigoCompetencia_Col)
        
        'Lê o Competencias pelo NomeReduzido
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134944 Then gError 134343

        Set objCTCompetencias = New ClassCTCompetencias
        
        objCTCompetencias.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
        objCentrodeTrabalho.colCompetencias.Add objCTCompetencias
    
    Next
    
    'Ir preenchendo a colecao no objCentrodeTrabalho com todas as linhas "existentes"
    'do grid CTMaquinas
    For iIndice = 1 To objGridItens.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CodigoItem_Col))) = 0 Then Exit For
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_CodigoItem_Col)

        'Lê as Maquinas pelo NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 134344

        Set objCTMaquinas = New ClassCTMaquinas
        
        objCTMaquinas.lNumIntDocMaq = objMaquinas.lNumIntDoc
        objCTMaquinas.iQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        'Localiza na coleção os turnos desta maquina
        For Each objCTMaqCol In gcolMaquinas
        
            'se encontrou...
            If objCTMaqCol.lNumIntDocMaq = objMaquinas.lNumIntDoc Then
            
                'direciona a coleção do obj para os dados do grid
                Set objCTMaquinas.colTurnos = objCTMaqCol.colTurnos
                Exit For
            
            End If
            
        Next
    
        objCentrodeTrabalho.colMaquinas.Add objCTMaquinas
    
    Next

    'Ir preenchendo a colecao no objCentrodeTrabalho com todas as linhas "existentes"
    'do grid Operadores
    For iIndice = 1 To objGridOperadores.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridOperadores.TextMatrix(iIndice, iGrid_CodigoTipoMO_Col))) = 0 Then Exit For
        
        Set objCTOperadores = New ClassCTOperadores

        objCTOperadores.iCodTipoMO = StrParaInt(GridOperadores.TextMatrix(iIndice, iGrid_CodigoTipoMO_Col))
        objCTOperadores.iQuantidade = StrParaDbl(GridOperadores.TextMatrix(iIndice, iGrid_QuantidadeMO_Col))
        
        objCentrodeTrabalho.colOperadores.Add objCTOperadores
    
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 134343, 134344

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144345)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CentrodeTrabalho"

    'Lê os dados da Tela Centro de Trabalho
    lErro = Move_Tela_Memoria(objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134345

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCentrodeTrabalho.lCodigo, 0, "Codigo"
    colCampoValor.Add "FilialEmpresa", giFilialEmpresa, 0, "FilialEmpresa"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 134345

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144346)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho

On Error GoTo Erro_Tela_Preenche

    objCentrodeTrabalho.lCodigo = colCampoValor.Item("Codigo").vValor

    objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa

    If Len(Trim(objCentrodeTrabalho.lCodigo)) > 0 And objCentrodeTrabalho.iFilialEmpresa <> 0 Then
        lErro = Traz_CentrodeTrabalho_Tela(objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 134346
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 134346

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144347)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 134347

    'Verifica se NomeReduzido está preenchido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 134348

    'Verifica se a Descrição está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 134548

    'Verifica se Turnos está preenchido
    If Len(Trim(QtdeTurnos.Text)) = 0 Then gError 134349

    'Verifica se HorasTurno está preenchido
    If Len(Trim(HorasTurno.Text)) = 0 Then gError 134350
    
    'Para cada CTMaquinas
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        'Verifica se a Quantidade foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 134351

    Next
    
    'Verifica se existem Competências cadastradas
    If objGridCompetencias.iLinhasExistentes = 0 Then gError 134352

    'Para cada CTOperadores
    For iIndice = 1 To objGridOperadores.iLinhasExistentes
        
        'Verifica se a Quantidade foi informada
        If Len(Trim(GridOperadores.TextMatrix(iIndice, iGrid_QuantidadeMO_Col))) = 0 Then gError 139107

    Next

    'Preenche o objCentrodeTrabalho
    lErro = Move_Tela_Memoria(objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134353
    
    lErro = Trata_Alteracao(objCentrodeTrabalho, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 137688

    'Grava o/a CentrodeTrabalho no Banco de Dados
    lErro = CF("CentrodeTrabalho_Grava", objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134354

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134347
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)

        Case 134348
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMERED_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)
        
        Case 134349
            Call Rotina_Erro(vbOKOnly, "ERRO_TURNOS_NAO_PREENCHIDO", gErr)
            
        Case 134350
            Call Rotina_Erro(vbOKOnly, "ERRO_HORASTURNOS_NAO_PREENCHIDO", gErr)
            
        Case 134351
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_MAQUINAS_NAO_PREENCHIDO", gErr)
            
        Case 139107
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_OPERADORES_NAO_PREENCHIDO", gErr)
            
        Case 134352
            Call Rotina_Erro(vbOKOnly, "ERRO_CTCOMPETENCIAS_NAO_PREENCHIDO", gErr)
            
        Case 134548
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
        
        Case 134353, 134354, 137688
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144348)

    End Select

    Exit Function

End Function

Function Limpa_Tela_CentrodeTrabalho() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CentrodeTrabalho
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    'Limpa os Grids
    Call Grid_Limpa(objGridCompetencias)
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridTurnosDias)
    Call Grid_Limpa(objGridDisponibilidadeMaquina)
    Call Grid_Limpa(objGridOperadores)
    
    'Zera a Coleção de TurnosMaquinas
    Set gcolMaquinas = New Collection
    
    FrameDispMaquina.Caption = "Disponibilidade em Horas"
    
    'Marca os Dias úteis
    Call Marca_Default_Dias

    iAlterado = 0
    iTurnoAlterado = 0

    Limpa_Tela_CentrodeTrabalho = SUCESSO

    Exit Function

Erro_Limpa_Tela_CentrodeTrabalho:

    Limpa_Tela_CentrodeTrabalho = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144349)

    End Select

    Exit Function

End Function

Function Traz_CentrodeTrabalho_Tela(objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim iDiaDaSemana As Integer
Dim objTurnosDias As New ClassTurno
Dim iLinha As Integer
Dim objCompetencias As ClassCompetencias
Dim objCTCompetencias As New ClassCTCompetencias
Dim objMaquinas As ClassMaquinas
Dim objCTMaquinas As New ClassCTMaquinas
Dim objCTOperadores As New ClassCTOperadores
Dim objTipoMO As ClassTiposDeMaodeObra

On Error GoTo Erro_Traz_CentrodeTrabalho_Tela

    'Lê o CentrodeTrabalho que está sendo Passado
    lErro = CF("CentrodeTrabalho_Le_Completo", objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 137212 Then gError 134356
    
    If lErro <> SUCESSO Then gError 134357

    'Limpa a Tela
    Call Limpa_Tela_CentrodeTrabalho

    Codigo.Text = objCentrodeTrabalho.lCodigo
    Descricao.Text = objCentrodeTrabalho.sDescricao
    NomeReduzido.Text = objCentrodeTrabalho.sNomeReduzido
    If objCentrodeTrabalho.dCargaMin <> 0 Then CargaMin.Text = CStr(objCentrodeTrabalho.dCargaMin * 100)
    If objCentrodeTrabalho.dCargaMax <> 0 Then CargaMax.Text = CStr(objCentrodeTrabalho.dCargaMax * 100)
    
    If objCentrodeTrabalho.iTurnos <> 0 Then QtdeTurnos.Text = CStr(objCentrodeTrabalho.iTurnos)
    If objCentrodeTrabalho.dHorasTurno <> 0 Then HorasTurno.Text = Formata_Estoque(objCentrodeTrabalho.dHorasTurno)

    For iDiaDaSemana = DOMINGO To SABADO
    
        If objCentrodeTrabalho.iDiaisUteis(iDiaDaSemana) = MARCADO Then
            Dias.Selected(iDiaDaSemana - 1) = True
        Else
            Dias.Selected(iDiaDaSemana - 1) = False
        End If
    
    Next

    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGridTurnosDias)
    
    For Each objTurnosDias In objCentrodeTrabalho.colTurnos
    
        'Insere no Grid Turnos Dias
        For iDiaDaSemana = DOMINGO To SABADO
        
            If objTurnosDias.dQtdHoras(iDiaDaSemana) <> 0 Then
                GridTurnosDias.TextMatrix(objTurnosDias.iTurno, iDiaDaSemana) = Formata_Estoque(objTurnosDias.dQtdHoras(iDiaDaSemana))
            End If
            
        Next
    
    Next
    
    If objCentrodeTrabalho.iTurnos > 0 Then
        objGridTurnosDias.iLinhasExistentes = objCentrodeTrabalho.iTurnos - 1
    Else
        objGridTurnosDias.iLinhasExistentes = objCentrodeTrabalho.colTurnos.Count
    End If
    
    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGridCompetencias)
    
    iLinha = 1
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objCTCompetencias.lNumIntDocCompet
        
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 134358
        
        'Insere no Grid Competencias
        GridCompetencias.TextMatrix(iLinha, iGrid_CodigoCompetencia_Col) = objCompetencias.sNomeReduzido
        GridCompetencias.TextMatrix(iLinha, iGrid_DescricaoComp_Col) = objCompetencias.sDescricao
    
        iLinha = iLinha + 1
    
    Next

    objGridCompetencias.iLinhasExistentes = objCentrodeTrabalho.colCompetencias.Count
        
    'Limpa os Grids antes de colocar algo nele
    Call Grid_Limpa(objGridItens)
    
    Call Grid_Limpa(objGridDisponibilidadeMaquina)
    Set gcolMaquinas = New Collection
    
    iLinha = 1
    
    'Exibe os dados da coleção de Máquinas na tela
    For Each objCTMaquinas In objCentrodeTrabalho.colMaquinas
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.lNumIntDoc = objCTMaquinas.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquinas)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 134359
        
        'Insere no Grid Maquinas
        GridItens.TextMatrix(iLinha, iGrid_CodigoItem_Col) = objMaquinas.sNomeReduzido
        GridItens.TextMatrix(iLinha, iGrid_DescricaoItem_Col) = objMaquinas.sDescricao
        GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = objCTMaquinas.iQuantidade
        
        'insere objCTMaquinas na coleção
        gcolMaquinas.Add objCTMaquinas
    
        iLinha = iLinha + 1
    
    Next

    objGridItens.iLinhasExistentes = objCentrodeTrabalho.colMaquinas.Count
    
    iLinha = 1
    
    'Exibe os dados da coleção de Operadores na tela
    For Each objCTOperadores In objCentrodeTrabalho.colOperadores
        
        Set objTipoMO = New ClassTiposDeMaodeObra
        
        objTipoMO.iCodigo = objCTOperadores.iCodTipoMO
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTipoMO)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 139086
        
        'Insere no Grid Operadores
        GridOperadores.TextMatrix(iLinha, iGrid_CodigoTipoMO_Col) = objTipoMO.iCodigo
        GridOperadores.TextMatrix(iLinha, iGrid_DescricaoMo_Col) = objTipoMO.sDescricao
        GridOperadores.TextMatrix(iLinha, iGrid_QuantidadeMO_Col) = objCTOperadores.iQuantidade
        
        iLinha = iLinha + 1
    
    Next

    objGridOperadores.iLinhasExistentes = objCentrodeTrabalho.colOperadores.Count
    
    iAlterado = 0
    iTurnoAlterado = 0
    
    Traz_CentrodeTrabalho_Tela = SUCESSO

    Exit Function

Erro_Traz_CentrodeTrabalho_Tela:

    Traz_CentrodeTrabalho_Tela = gErr

    Select Case gErr

        Case 134356, 134358, 134359, 139086
            'erros tratados nas rotinas chamadas
            
        Case 134357 'Sem dados - tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144350)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 134360

    'Limpa Tela
    Call Limpa_Tela_CentrodeTrabalho

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 134360

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144351)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144352)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134361

    Call Limpa_Tela_CentrodeTrabalho

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 134361

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144353)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 134362

    objCentrodeTrabalho.lCodigo = StrParaLong(Codigo.Text)
    objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CENTRODETRABALHO", objCentrodeTrabalho.lCodigo)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui o Centro de Trabaho
    lErro = CF("CentrodeTrabalho_Exclui", objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134363

    'Limpa Tela
    Call Limpa_Tela_CentrodeTrabalho

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134362
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)

        Case 134363
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144354)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Veifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

        'Critica a Codigo
        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 134364

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 134364
            'erro tratados na rotinas chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144355)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CargaMin_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CargaMin_Validate

    'Veifica se CargaMin está preenchida
    If Len(Trim(CargaMin.Text)) <> 0 Then

       'Critica a CargaMin
       lErro = Porcentagem_Critica(CargaMin.Text)
       If lErro <> SUCESSO Then gError 134367

    End If

    Exit Sub

Erro_CargaMin_Validate:

    Cancel = True

    Select Case gErr

        Case 134367

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144356)

    End Select

    Exit Sub

End Sub

Private Sub CargaMin_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CargaMin, iAlterado)
    
End Sub

Private Sub CargaMin_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CargaMax_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CargaMax_Validate

    'Veifica se CargaMax está preenchida
    If Len(Trim(CargaMax.Text)) <> 0 Then

       'Critica a CargaMax
       lErro = Porcentagem_Critica(CargaMax.Text)
       If lErro <> SUCESSO Then gError 134368

    End If

    Exit Sub

Erro_CargaMax_Validate:

    Cancel = True

    Select Case gErr

        Case 134368

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144357)

    End Select

    Exit Sub

End Sub

Private Sub CargaMax_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CargaMax, iAlterado)
    
End Sub

Private Sub CargaMax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCentrodeTrabalho = obj1

    'Mostra os dados do CentrodeTrabalho na tela
    lErro = Traz_CentrodeTrabalho_Tela(objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134369
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 134369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144358)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objCentrodeTrabalho.lCodigo = StrParaLong(Codigo.Text)

    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144359)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_LostFocus()

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridCompetencias_LostFocus()

    Call Grid_Libera_Foco(objGridCompetencias)

End Sub

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Máquina")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Quantidade")

    'Controles que participam do Grid
    objGrid.colCampo.Add (CodigoItem.Name)
    objGrid.colCampo.Add (DescricaoItem.Name)
    objGrid.colCampo.Add (Quantidade.Name)

    'Colunas do Grid
    iGrid_CodigoItem_Col = 1
    iGrid_DescricaoItem_Col = 2
    iGrid_Quantidade_Col = 3

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridItens, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridItens, iAlterado)
        End If

End Sub

Private Sub GridItens_GotFocus()
    
    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_LeaveCell()
    
    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()

Dim lErro As Long

On Error GoTo Erro_GridItens_RowColChange

    Call Grid_RowColChange(objGridItens)
    
    If (GridItens.Row <> iLinhaAntiga) Then

        lErro = Preenche_GridDispMaquina()
        If lErro <> SUCESSO Then gError 137293

        'Guarda a Linha corrente
        iLinhaAntiga = GridItens.Row

    End If

    Exit Sub

Erro_GridItens_RowColChange:

    Select Case gErr

        Case 137293
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144360)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CodigoItem_Col
                
                    lErro = Saida_Celula_CodigoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 134370
    
                Case iGrid_DescricaoItem_Col
    
                    lErro = Saida_Celula_DescricaoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 134371
        
                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
        
            End Select
        
        'Competencias
        ElseIf objGridInt.objGrid.Name = GridCompetencias.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CodigoCompetencia_Col
                
                    lErro = Saida_Celula_CodigoCompetencia(objGridInt)
                    If lErro <> SUCESSO Then gError 134373
                    
                Case iGrid_DescricaoComp_Col
                
                    lErro = Saida_Celula_DescricaoComp(objGridInt)
                    If lErro <> SUCESSO Then gError 134374
                    
            End Select
        
        'TurnosDias
        ElseIf objGridInt.objGrid.Name = GridTurnosDias.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_DispCTDom_Col

                    lErro = Saida_Celula_DispCTDom(objGridInt)
                    If lErro <> SUCESSO Then gError 137293

                Case iGrid_DispCTSeg_Col

                    lErro = Saida_Celula_DispCTSeg(objGridInt)
                    If lErro <> SUCESSO Then gError 137294

                Case iGrid_DispCTTer_Col

                    lErro = Saida_Celula_DispCTTer(objGridInt)
                    If lErro <> SUCESSO Then gError 137295

                Case iGrid_DispCTQua_Col

                    lErro = Saida_Celula_DispCTQua(objGridInt)
                    If lErro <> SUCESSO Then gError 137296

                Case iGrid_DispCTQui_Col

                    lErro = Saida_Celula_DispCTQui(objGridInt)
                    If lErro <> SUCESSO Then gError 137297

                Case iGrid_DispCTSex_Col

                    lErro = Saida_Celula_DispCTSex(objGridInt)
                    If lErro <> SUCESSO Then gError 137298

                Case iGrid_DispCTSab_Col

                    lErro = Saida_Celula_DispCTSab(objGridInt)
                    If lErro <> SUCESSO Then gError 137299

            End Select
                    
        'DisponibilidadeMaquina
        ElseIf objGridInt.objGrid.Name = GridDisponibilidadeMaquina.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_TurnoMaq_Col

                    lErro = Saida_Celula_TurnoMaq(objGridInt)
                    If lErro <> SUCESSO Then gError 137300

                Case iGrid_DispMaqDom_Col

                    lErro = Saida_Celula_DispMaqDom(objGridInt)
                    If lErro <> SUCESSO Then gError 137301

                Case iGrid_DispMaqSeg_Col

                    lErro = Saida_Celula_DispMaqSeg(objGridInt)
                    If lErro <> SUCESSO Then gError 137302

                Case iGrid_DispMaqTer_Col

                    lErro = Saida_Celula_DispMaqTer(objGridInt)
                    If lErro <> SUCESSO Then gError 137303

                Case iGrid_DispMaqQua_Col

                    lErro = Saida_Celula_DispMaqQua(objGridInt)
                    If lErro <> SUCESSO Then gError 137304

                Case iGrid_DispMaqQui_Col

                    lErro = Saida_Celula_DispMaqQui(objGridInt)
                    If lErro <> SUCESSO Then gError 137305

                Case iGrid_DispMaqSex_Col

                    lErro = Saida_Celula_DispMaqSex(objGridInt)
                    If lErro <> SUCESSO Then gError 137306

                Case iGrid_DispMaqSab_Col

                    lErro = Saida_Celula_DispMaqSab(objGridInt)
                    If lErro <> SUCESSO Then gError 137307

            End Select
                
        'Verifica se é o GridOperadores
        ElseIf objGridInt.objGrid.Name = GridOperadores.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CodigoTipoMO_Col
                
                    lErro = Saida_Celula_CodTipoMO(objGridInt)
                    If lErro <> SUCESSO Then gError 139076
    
                Case iGrid_QuantidadeMO_Col
    
                    lErro = Saida_Celula_QuantidadeMO(objGridInt)
                    If lErro <> SUCESSO Then gError 139077
        
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 134375

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 134370 To 134374, 137293 To 137307, 139076, 139077
            'erros tratatos nas rotinas chamadas
        
        Case 134375
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144361)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridCompetencias(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Competência")
    objGrid.colColuna.Add ("Descricao")

    'Controles que participam do Grid
    objGrid.colCampo.Add (CodigoCompetencia.Name)
    objGrid.colCampo.Add (DescricaoComp.Name)

    'Colunas do Grid
    iGrid_CodigoCompetencia_Col = 1
    iGrid_DescricaoComp_Col = 2

    objGrid.objGrid = GridCompetencias

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridCompetencias.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridCompetencias = SUCESSO

End Function

Private Sub GridCompetencias_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridCompetencias, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridCompetencias, iAlterado)
        End If

End Sub

Private Sub GridCompetencias_GotFocus()
    
    Call Grid_Recebe_Foco(objGridCompetencias)

End Sub

Private Sub GridCompetencias_EnterCell()

    Call Grid_Entrada_Celula(objGridCompetencias, iAlterado)

End Sub

Private Sub GridCompetencias_LeaveCell()
    
    Call Saida_Celula(objGridCompetencias)

End Sub

Private Sub GridCompetencias_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridCompetencias, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCompetencias, iAlterado)
    End If

End Sub

Private Sub GridCompetencias_RowColChange()

    Call Grid_RowColChange(objGridCompetencias)

End Sub

Private Sub GridCompetencias_Scroll()

    Call Grid_Scroll(objGridCompetencias)

End Sub

Private Function Saida_Celula_CodigoCompetencia(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CodigoCompetencia do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim objCompetencias As ClassCompetencias

On Error GoTo Erro_Saida_Celula_CodigoCompetencia

    Set objGridInt.objControle = CodigoCompetencia

    'Se o campo foi preenchido
    If Len(Trim(CodigoCompetencia.Text)) > 0 Then
                
        Set objCompetencias = New ClassCompetencias
        
        'Verifica sua existencia
        lErro = CF("TP_Competencia_Le", CodigoCompetencia, objCompetencias)
        If lErro <> SUCESSO Then gError 134377
        
        'Verifica se há alguma competencia repetida no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridCompetencias.Row Then
                                                    
                If GridCompetencias.TextMatrix(iLinha, iGrid_CodigoCompetencia_Col) = objCompetencias.sNomeReduzido Then
                
                    CodigoCompetencia.Text = ""
                    gError 134376
                    
                End If
                    
            End If
                           
        Next
                
        GridCompetencias.TextMatrix(GridCompetencias.Row, iGrid_DescricaoComp_Col) = objCompetencias.sDescricao
            
        'verifica se precisa preencher o grid com uma nova linha
        If GridCompetencias.Row - GridCompetencias.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134379

    Saida_Celula_CodigoCompetencia = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigoCompetencia:

    Saida_Celula_CodigoCompetencia = gErr

    Select Case gErr
    
        Case 134376
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_REPETIDA", gErr, objCompetencias.sNomeReduzido, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 134377, 134379
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144362)

    End Select

    Exit Function

End Function

Private Sub BotaoCompetencia_Click()

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim colSelecao As New Collection

On Error GoTo Erro_botaoCompetencia_Click

    Set objCompetencias = New ClassCompetencias

    If Me.ActiveControl Is CodigoCompetencia Then
    
        objCompetencias.sNomeReduzido = CodigoCompetencia.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridCompetencias.Row = 0 Then gError 134380

        objCompetencias.sNomeReduzido = GridCompetencias.TextMatrix(GridCompetencias.Row, iGrid_CodigoCompetencia_Col)
        
    End If
    
    'Verifica a Competencia no BD a partir do NomeReduzido
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 137923
        
    Call Chama_Tela("CompetenciasLista", colSelecao, objCompetencias, objEventoCompetencia)

    Exit Sub

Erro_botaoCompetencia_Click:

    Select Case gErr

        Case 134380
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 137923

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144363)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoMaquinas_Click()

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMaquinas_Click

    Set objMaquinas = New ClassMaquinas

    If Me.ActiveControl Is CodigoItem Then
            
        objMaquinas.sNomeReduzido = CodigoItem.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 134381

        objMaquinas.sNomeReduzido = GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)
        
    End If
    
    'Le a Máquina no BD a partir do NomeReduzido
    lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
    If lErro <> SUCESSO And lErro <> 103100 Then gError 137930
    
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaquina)

    Exit Sub

Erro_BotaoMaquinas_Click:

    Select Case gErr

        Case 134381
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 137930

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144364)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCompetencia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim iLinha As Integer

On Error GoTo Erro_objEventoCompetencia_evSelecao

    Set objCompetencias = obj1
        
    'Verifica sua existencia
    lErro = CF("TP_Competencia_Le", CodigoCompetencia, objCompetencias)
    If lErro <> SUCESSO Then gError 134382
    
    'Verifica se há alguma competencia repetida no grid
    For iLinha = 1 To objGridCompetencias.iLinhasExistentes
        
        If iLinha <> GridCompetencias.Row Then
                                                
            If GridCompetencias.TextMatrix(iLinha, iGrid_CodigoCompetencia_Col) = objCompetencias.sNomeReduzido Then
            
                CodigoCompetencia.Text = ""
                gError 134384
                
            End If
                
        End If
                       
    Next
    
    CodigoCompetencia.Text = objCompetencias.sNomeReduzido
    
    If Not (Me.ActiveControl Is CodigoCompetencia) Then
        GridCompetencias.TextMatrix(GridCompetencias.Row, iGrid_CodigoCompetencia_Col) = objCompetencias.sNomeReduzido
        GridCompetencias.TextMatrix(GridCompetencias.Row, iGrid_DescricaoComp_Col) = objCompetencias.sDescricao
    End If

    'verifica se precisa preencher o grid com uma nova linha
    If GridCompetencias.Row - GridCompetencias.FixedRows = objGridCompetencias.iLinhasExistentes Then
        objGridCompetencias.iLinhasExistentes = objGridCompetencias.iLinhasExistentes + 1
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoCompetencia_evSelecao:

    Select Case gErr

        Case 134382
            'erro tratado na rotina chamada
        
        Case 134384
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_REPETIDA", gErr, objCompetencias.sNomeReduzido, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144365)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMaquina_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim iLinha As Integer

On Error GoTo Erro_objEventoMaquina_evSelecao

    Set objMaquinas = obj1

    'Lê o Maquinas
    lErro = CF("TP_Maquina_Le", CodigoItem, objMaquinas)
    If lErro <> SUCESSO Then gError 134385
            
    For iLinha = 1 To objGridItens.iLinhasExistentes
        
        If iLinha <> GridItens.Row Then
                                                
            If GridItens.TextMatrix(iLinha, iGrid_CodigoItem_Col) = objMaquinas.sNomeReduzido Then
            
                CodigoItem.Text = ""
                gError 134387
                
            End If
                
        End If
                       
    Next
    
    'Mostra os dados do Maquinas na tela
    CodigoItem.Text = objMaquinas.sNomeReduzido
    
    If Not (Me.ActiveControl Is CodigoItem) Then
        GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col) = objMaquinas.sNomeReduzido
        GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoItem_Col) = objMaquinas.sDescricao
    End If
    
    lErro = Preenche_GridDispMaquina()
    If lErro <> SUCESSO Then gError 137999
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoMaquina_evSelecao:

    Select Case gErr

        Case 134385
            'erro tratado na rotina chamada
        
        Case 134387
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_REPETIDA", gErr, objMaquinas.sNomeReduzido, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144366)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_CodigoItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CodigoItem do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Saida_Celula_CodigoItem

    Set objGridInt.objControle = CodigoItem

    'Se o campo foi preenchido
    If Len(Trim(CodigoItem.Text)) > 0 Then
    
        Set objMaquinas = New ClassMaquinas
    
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", CodigoItem, objMaquinas)
        If lErro <> SUCESSO Then gError 134389
    
        'Verifica se há alguma maquina repetida no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridItens.Row Then
                                                    
                If GridItens.TextMatrix(iLinha, iGrid_CodigoItem_Col) = objMaquinas.sNomeReduzido Then
                
                    CodigoItem.Text = ""
                    gError 134388
                    
                End If
                    
            End If
                           
        Next
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoItem_Col) = objMaquinas.sDescricao
        
        lErro = Preenche_GridDispMaquina()
        If lErro <> SUCESSO Then gError 137308

        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134391

    Saida_Celula_CodigoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigoItem:

    Saida_Celula_CodigoItem = gErr

    Select Case gErr
        
        Case 134388
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_REPETIDA", gErr, objMaquinas.sNomeReduzido, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134389, 134391, 137308
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144367)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CodigoItem do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    'Se o campo foi preenchido
    If Len(Trim(DescricaoItem.Text)) > 0 Then
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134392

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = gErr

    Select Case gErr
        
        Case 134392
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144368)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'Se o campo foi preenchido
    If Len(Trim(Quantidade.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 134393

        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144369)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sCodItem As String
Dim sCodComp As String
Dim iCodTipoMO As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    If GridItens.Row > 0 Then
        'Guardo o valor do Codigo do Item
        sCodItem = GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)
    End If

    If GridCompetencias.Row > 0 Then
        'Guardo o valor do Codigo da Competência
        sCodComp = GridCompetencias.TextMatrix(GridCompetencias.Row, iGrid_CodigoCompetencia_Col)
    End If
    
    If GridOperadores.Row > 0 Then
        'Guardo o valor do Codigo da Mão de Obra
        iCodTipoMO = StrParaInt(GridOperadores.TextMatrix(GridOperadores.Row, iGrid_CodigoTipoMO_Col))
    End If
    
    Select Case objControl.Name
    
        'Grid Itens
        Case Is = "CodigoItem"
            
            If Len(sCodItem) > 0 Then
                objControl.Enabled = False
    
            Else
                objControl.Enabled = True
            
            End If
    
        Case Is = "DescricaoItem"
    
            objControl.Enabled = False
            
        Case Is = "Quantidade"
            
            If Len(sCodItem) > 0 Then
                objControl.Enabled = True
    
            Else
                objControl.Enabled = False
            
            End If
            
        'Grid Competencias
        Case Is = "CodigoCompetencia"
            
            If Len(sCodComp) > 0 Then
                objControl.Enabled = False
    
            Else
                objControl.Enabled = True
            
            End If
        
        Case Is = "DescricaoComp"
    
            objControl.Enabled = False
        
        'Grid TurnosDias
        Case Is = "DispCTDom"
            
            objControl.Enabled = Dias.Selected(DOMINGO - 1)
    
        Case Is = "DispCTSeg"
        
            objControl.Enabled = Dias.Selected(SEGUNDA - 1)
    
        Case Is = "DispCTTer"
        
            objControl.Enabled = Dias.Selected(TERCA - 1)
    
        Case Is = "DispCTQua"
        
            objControl.Enabled = Dias.Selected(QUARTA - 1)
    
        Case Is = "DispCTQui"
        
            objControl.Enabled = Dias.Selected(QUINTA - 1)
    
        Case Is = "DispCTSex"
        
            objControl.Enabled = Dias.Selected(SEXTA - 1)
    
        Case Is = "DispCTSab"
        
            objControl.Enabled = Dias.Selected(SABADO - 1)
        
        'Grid DisponibilidadeMaquina
        Case Is = "TurnoMaq"
        
            'se Turno preenchido
            If Len(GridDisponibilidadeMaquina.TextMatrix(GridDisponibilidadeMaquina.Row, iGrid_TurnoMaq_Col)) <> 0 Then
                
                objControl.Enabled = False
                
            Else
            
                objControl.Enabled = True
            
            End If
        
        Case Is = "DispMaqDom", "DispMaqSeg", "DispMaqTer", "DispMaqQua", "DispMaqQui", "DispMaqSex", "DispMaqSab"
                                
            'se Turno preenchido
            If Len(GridDisponibilidadeMaquina.TextMatrix(GridDisponibilidadeMaquina.Row, iGrid_TurnoMaq_Col)) <> 0 Then
                
                objControl.Enabled = True
                
            Else
            
                objControl.Enabled = False
            
            End If
            
        'Grid Operadores
        Case Is = "CodTipoMO"
            
            If iCodTipoMO > 0 Then
                objControl.Enabled = False
    
            Else
                objControl.Enabled = True
            
            End If
    
        Case Is = "DescricaoMO"
    
            objControl.Enabled = False
            
        Case Is = "Quantidade"
            
            If iCodTipoMO > 0 Then
                objControl.Enabled = True
    
            Else
                objControl.Enabled = False
            
            End If
                        
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144370)

    End Select

    Exit Sub

End Sub

Private Sub CodigoCompetencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoCompetencia_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCompetencias)

End Sub

Private Sub CodigoCompetencia_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCompetencias)

End Sub

Private Sub CodigoCompetencia_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCompetencias.objControle = CodigoCompetencia
    lErro = Grid_Campo_Libera_Foco(objGridCompetencias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CodigoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CodigoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CodigoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CodigoItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoComp_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoComp_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCompetencias)

End Sub

Private Sub DescricaoComp_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCompetencias)

End Sub

Private Sub DescricaoComp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCompetencias.objControle = DescricaoComp
    lErro = Grid_Campo_Libera_Foco(objGridCompetencias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_DescricaoComp(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DescricaoComp do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_DescricaoComp

    Set objGridInt.objControle = DescricaoComp

    'Se o campo foi preenchido
    If Len(Trim(DescricaoComp.Text)) > 0 Then
                                
        'verifica se precisa preencher o grid com uma nova linha
        If GridCompetencias.Row - GridCompetencias.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134397

    Saida_Celula_DescricaoComp = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoComp:

    Saida_Celula_DescricaoComp = gErr

    Select Case gErr
        
        Case 134397
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144371)

    End Select

    Exit Function

End Function

'#################################################

Private Function Inicializa_GridTurnosDias(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("Turno")
    objGrid.colColuna.Add ("Domingo")
    objGrid.colColuna.Add ("Segunda")
    objGrid.colColuna.Add ("Terça")
    objGrid.colColuna.Add ("Quarta")
    objGrid.colColuna.Add ("Quinta")
    objGrid.colColuna.Add ("Sexta")
    objGrid.colColuna.Add ("Sabado")

    'Controles que participam do Grid
    objGrid.colCampo.Add (DispCTDom.Name)
    objGrid.colCampo.Add (DispCTSeg.Name)
    objGrid.colCampo.Add (DispCTTer.Name)
    objGrid.colCampo.Add (DispCTQua.Name)
    objGrid.colCampo.Add (DispCTQui.Name)
    objGrid.colCampo.Add (DispCTSex.Name)
    objGrid.colCampo.Add (DispCTSab.Name)

    'Colunas do Grid
    iGrid_DispCTDom_Col = 1
    iGrid_DispCTSeg_Col = 2
    iGrid_DispCTTer_Col = 3
    iGrid_DispCTQua_Col = 4
    iGrid_DispCTQui_Col = 5
    iGrid_DispCTSex_Col = 6
    iGrid_DispCTSab_Col = 7

    objGrid.objGrid = GridTurnosDias

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridTurnosDias.ColWidth(0) = 600

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridTurnosDias = SUCESSO

End Function

Private Sub GridTurnosDias_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridTurnosDias, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridTurnosDias, iAlterado)
        End If

End Sub

Private Sub GridTurnosDias_GotFocus()

    Call Grid_Recebe_Foco(objGridTurnosDias)

End Sub

Private Sub GridTurnosDias_EnterCell()

    Call Grid_Entrada_Celula(objGridTurnosDias, iAlterado)

End Sub

Private Sub GridTurnosDias_LeaveCell()

    Call Saida_Celula(objGridTurnosDias)

End Sub

Private Sub GridTurnosDias_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridTurnosDias, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTurnosDias, iAlterado)
    End If

End Sub

Private Sub GridTurnosDias_RowColChange()

    Call Grid_RowColChange(objGridTurnosDias)

End Sub

Private Sub GridTurnosDias_Scroll()

    Call Grid_Scroll(objGridTurnosDias)

End Sub

Private Sub DispCTDom_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispCTDom_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnosDias)

End Sub

Private Sub DispCTDom_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnosDias)

End Sub

Private Sub DispCTDom_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnosDias.objControle = DispCTDom
    lErro = Grid_Campo_Libera_Foco(objGridTurnosDias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispCTSeg_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispCTSeg_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnosDias)

End Sub

Private Sub DispCTSeg_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnosDias)

End Sub

Private Sub DispCTSeg_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnosDias.objControle = DispCTSeg
    lErro = Grid_Campo_Libera_Foco(objGridTurnosDias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispCTTer_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispCTTer_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnosDias)

End Sub

Private Sub DispCTTer_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnosDias)

End Sub

Private Sub DispCTTer_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnosDias.objControle = DispCTTer
    lErro = Grid_Campo_Libera_Foco(objGridTurnosDias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispCTQua_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispCTQua_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnosDias)

End Sub

Private Sub DispCTQua_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnosDias)

End Sub

Private Sub DispCTQua_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnosDias.objControle = DispCTQua
    lErro = Grid_Campo_Libera_Foco(objGridTurnosDias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispCTQui_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispCTQui_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnosDias)

End Sub

Private Sub DispCTQui_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnosDias)

End Sub

Private Sub DispCTQui_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnosDias.objControle = DispCTQui
    lErro = Grid_Campo_Libera_Foco(objGridTurnosDias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispCTSex_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispCTSex_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnosDias)

End Sub

Private Sub DispCTSex_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnosDias)

End Sub

Private Sub DispCTSex_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnosDias.objControle = DispCTSex
    lErro = Grid_Campo_Libera_Foco(objGridTurnosDias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispCTSab_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispCTSab_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnosDias)

End Sub

Private Sub DispCTSab_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnosDias)

End Sub

Private Sub DispCTSab_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnosDias.objControle = DispCTSab
    lErro = Grid_Campo_Libera_Foco(objGridTurnosDias)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Inicializa_GridDispMaquina(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Turno")
    objGrid.colColuna.Add ("Domingo")
    objGrid.colColuna.Add ("Segunda")
    objGrid.colColuna.Add ("Terça")
    objGrid.colColuna.Add ("Quarta")
    objGrid.colColuna.Add ("Quinta")
    objGrid.colColuna.Add ("Sexta")
    objGrid.colColuna.Add ("Sabado")

    'Controles que participam do Grid
    objGrid.colCampo.Add (TurnoMaq.Name)
    objGrid.colCampo.Add (DispMaqDom.Name)
    objGrid.colCampo.Add (DispMaqSeg.Name)
    objGrid.colCampo.Add (DispMaqTer.Name)
    objGrid.colCampo.Add (DispMaqQua.Name)
    objGrid.colCampo.Add (DispMaqQui.Name)
    objGrid.colCampo.Add (DispMaqSex.Name)
    objGrid.colCampo.Add (DispMaqSab.Name)

    'Colunas do Grid
    iGrid_TurnoMaq_Col = 1
    iGrid_DispMaqDom_Col = 2
    iGrid_DispMaqSeg_Col = 3
    iGrid_DispMaqTer_Col = 4
    iGrid_DispMaqQua_Col = 5
    iGrid_DispMaqQui_Col = 6
    iGrid_DispMaqSex_Col = 7
    iGrid_DispMaqSab_Col = 8

    objGrid.objGrid = GridDisponibilidadeMaquina

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1
    
    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridDisponibilidadeMaquina.ColWidth(0) = 600

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridDispMaquina = SUCESSO

End Function

Private Sub GridDisponibilidadeMaquina_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridDisponibilidadeMaquina, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridDisponibilidadeMaquina, iAlterado)
        End If

End Sub

Private Sub GridDisponibilidadeMaquina_GotFocus()

    Call Grid_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub

Private Sub GridDisponibilidadeMaquina_EnterCell()

    Call Grid_Entrada_Celula(objGridDisponibilidadeMaquina, iAlterado)

End Sub

Private Sub GridDisponibilidadeMaquina_LeaveCell()

    Call Saida_Celula(objGridDisponibilidadeMaquina)

End Sub

Private Sub GridDisponibilidadeMaquina_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDisponibilidadeMaquina, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDisponibilidadeMaquina, iAlterado)
    End If

End Sub

Private Sub GridDisponibilidadeMaquina_RowColChange()

    Call Grid_RowColChange(objGridDisponibilidadeMaquina)

End Sub

Private Sub GridDisponibilidadeMaquina_Scroll()

    Call Grid_Scroll(objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqDom_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispMaqDom_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqDom_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqDom_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDisponibilidadeMaquina.objControle = DispMaqDom
    lErro = Grid_Campo_Libera_Foco(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispMaqSeg_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispMaqSeg_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqSeg_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqSeg_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDisponibilidadeMaquina.objControle = DispMaqSeg
    lErro = Grid_Campo_Libera_Foco(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispMaqTer_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispMaqTer_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqTer_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqTer_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDisponibilidadeMaquina.objControle = DispMaqTer
    lErro = Grid_Campo_Libera_Foco(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispMaqQua_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispMaqQua_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqQua_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqQua_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDisponibilidadeMaquina.objControle = DispMaqQua
    lErro = Grid_Campo_Libera_Foco(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispMaqQui_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispMaqQui_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqQui_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqQui_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDisponibilidadeMaquina.objControle = DispMaqQui
    lErro = Grid_Campo_Libera_Foco(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispMaqSex_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispMaqSex_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqSex_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqSex_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDisponibilidadeMaquina.objControle = DispMaqSex
    lErro = Grid_Campo_Libera_Foco(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DispMaqSab_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DispMaqSab_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqSab_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDisponibilidadeMaquina)

End Sub

Private Sub DispMaqSab_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDisponibilidadeMaquina.objControle = DispMaqSab
    lErro = Grid_Campo_Libera_Foco(objGridDisponibilidadeMaquina)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_DispCTDom(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispCTDom

    Set objGridInt.objControle = DispCTDom

    'Se o campo foi preenchido
    If Len(DispCTDom.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispCTDom.Text, GridTurnosDias.Row, iGrid_DispCTDom_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137309
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137310

    Saida_Celula_DispCTDom = SUCESSO

    Exit Function

Erro_Saida_Celula_DispCTDom:

    Saida_Celula_DispCTDom = gErr

    Select Case gErr

        Case 137309, 137310
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144372)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TurnoMaq(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objTurno As ClassTurno
Dim iDiaDaSemana As Integer
Dim dHorasGridDias As Double
Dim iLinha As Integer
Dim sTurno As String

On Error GoTo Erro_Saida_Celula_TurnoMaq

    Set objGridInt.objControle = TurnoMaq

    'Se o campo foi preenchido
    If Len(TurnoMaq.Text) > 0 Then

        If GridItens.Row = 0 Then gError 140856

        'Critica o Turno
        lErro = Inteiro_Critica(TurnoMaq.Text)
        If lErro <> SUCESSO Then gError 137311
        
        'Verifica se o turno está repetido no grid
        For iLinha = 1 To objGridDisponibilidadeMaquina.iLinhasExistentes
            
            If iLinha <> GridDisponibilidadeMaquina.Row Then
                                                    
                If GridDisponibilidadeMaquina.TextMatrix(iLinha, iGrid_TurnoMaq_Col) = TurnoMaq.Text Then
                    
                    sTurno = TurnoMaq.Text
                    TurnoMaq.Text = ""
                    gError 137892
                    
                End If
                    
            End If
                           
        Next
        
        Set objTurno = New ClassTurno
        
        'com os dados passados
        objTurno.iTurno = StrParaInt(TurnoMaq.Text)
    
        'se tem o turno no grid TurnosDias
        If objTurno.iTurno > 0 And objTurno.iTurno <= StrParaInt(QtdeTurnos.Text) Then
    
            'e para cada dia da Semana
            For iDiaDaSemana = DOMINGO To SABADO
            
                'se for Dia Útil...
                If Dias.Selected(iDiaDaSemana - 1) = True Then
                
                    If Len(GridTurnosDias.TextMatrix(objTurno.iTurno, iDiaDaSemana)) <> 0 Then
                    
                        dHorasGridDias = StrParaDbl(GridTurnosDias.TextMatrix(objTurno.iTurno, iDiaDaSemana))
                
                        'Preenche com a quantidade default do GridTurnoDias
                        GridDisponibilidadeMaquina.TextMatrix(GridDisponibilidadeMaquina.Row, iDiaDaSemana + 1) = Formata_Estoque(dHorasGridDias)
                    
                        objTurno.dQtdHoras(iDiaDaSemana) = dHorasGridDias
                        
                    End If
                
                End If
            
            Next
        
        End If
        
        'e adiciona o objeto a coleção
        gcolMaquinas.Item(GridItens.Row).colTurnos.Add objTurno
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridDisponibilidadeMaquina.Row - GridDisponibilidadeMaquina.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137313

    Saida_Celula_TurnoMaq = SUCESSO

    Exit Function

Erro_Saida_Celula_TurnoMaq:

    Saida_Celula_TurnoMaq = gErr

    Select Case gErr

        Case 137311 To 137313
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 137892
            Call Rotina_Erro(vbOKOnly, "ERRO_TURNO_REPETIDO", gErr, sTurno, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 140856
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRIDMAQUINA_NAO_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144373)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispMaqDom(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispMaqDom

    Set objGridInt.objControle = DispMaqDom

    'Se o campo foi preenchido
    If Len(DispMaqDom.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispMaqDom.Text, GridDisponibilidadeMaquina.Row, iGrid_DispMaqDom_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137314
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridDisponibilidadeMaquina.Row - GridDisponibilidadeMaquina.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    gcolMaquinas.Item(GridItens.Row).colTurnos.Item(GridDisponibilidadeMaquina.Row).dQtdHoras(iGrid_DispMaqDom_Col - 1) = StrParaDbl(DispMaqDom.Text)
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137316

    Saida_Celula_DispMaqDom = SUCESSO

    Exit Function

Erro_Saida_Celula_DispMaqDom:

    Saida_Celula_DispMaqDom = gErr

    Select Case gErr

        Case 137314 To 137316
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144374)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispCTSeg(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispCTSeg

    Set objGridInt.objControle = DispCTSeg

    'Se o campo foi preenchido
    If Len(DispCTSeg.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispCTSeg.Text, GridTurnosDias.Row, iGrid_DispCTSeg_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137317
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137318

    Saida_Celula_DispCTSeg = SUCESSO

    Exit Function

Erro_Saida_Celula_DispCTSeg:

    Saida_Celula_DispCTSeg = gErr

    Select Case gErr

        Case 137317, 137318
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144375)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispMaqSeg(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispMaqSeg

    Set objGridInt.objControle = DispMaqSeg

    'Se o campo foi preenchido
    If Len(DispMaqSeg.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispMaqSeg.Text, GridDisponibilidadeMaquina.Row, iGrid_DispMaqSeg_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137319
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridDisponibilidadeMaquina.Row - GridDisponibilidadeMaquina.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    gcolMaquinas.Item(GridItens.Row).colTurnos.Item(GridDisponibilidadeMaquina.Row).dQtdHoras(iGrid_DispMaqSeg_Col - 1) = StrParaDbl(DispMaqSeg.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137321

    Saida_Celula_DispMaqSeg = SUCESSO

    Exit Function

Erro_Saida_Celula_DispMaqSeg:

    Saida_Celula_DispMaqSeg = gErr

    Select Case gErr

        Case 137319 To 137321
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144376)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispCTTer(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispCTTer

    Set objGridInt.objControle = DispCTTer

    'Se o campo foi preenchido
    If Len(DispCTTer.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispCTTer.Text, GridTurnosDias.Row, iGrid_DispCTTer_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137322
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137323

    Saida_Celula_DispCTTer = SUCESSO

    Exit Function

Erro_Saida_Celula_DispCTTer:

    Saida_Celula_DispCTTer = gErr

    Select Case gErr

        Case 137322, 137323
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144377)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispMaqTer(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispMaqTer

    Set objGridInt.objControle = DispMaqTer

    'Se o campo foi preenchido
    If Len(DispMaqTer.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispMaqTer.Text, GridDisponibilidadeMaquina.Row, iGrid_DispMaqTer_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137324
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridDisponibilidadeMaquina.Row - GridDisponibilidadeMaquina.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    gcolMaquinas.Item(GridItens.Row).colTurnos.Item(GridDisponibilidadeMaquina.Row).dQtdHoras(iGrid_DispMaqTer_Col - 1) = StrParaDbl(DispMaqTer.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137326

    Saida_Celula_DispMaqTer = SUCESSO

    Exit Function

Erro_Saida_Celula_DispMaqTer:

    Saida_Celula_DispMaqTer = gErr

    Select Case gErr

        Case 137324 To 137326
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144378)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispCTQua(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispCTQua

    Set objGridInt.objControle = DispCTQua

    'Se o campo foi preenchido
    If Len(DispCTQua.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispCTQua.Text, GridTurnosDias.Row, iGrid_DispCTQua_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137327
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137328

    Saida_Celula_DispCTQua = SUCESSO

    Exit Function

Erro_Saida_Celula_DispCTQua:

    Saida_Celula_DispCTQua = gErr

    Select Case gErr

        Case 137327, 137328
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144379)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispMaqQua(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispMaqQua

    Set objGridInt.objControle = DispMaqQua

    'Se o campo foi preenchido
    If Len(DispMaqQua.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispMaqQua.Text, GridDisponibilidadeMaquina.Row, iGrid_DispMaqQua_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137329
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridDisponibilidadeMaquina.Row - GridDisponibilidadeMaquina.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    gcolMaquinas.Item(GridItens.Row).colTurnos.Item(GridDisponibilidadeMaquina.Row).dQtdHoras(iGrid_DispMaqQua_Col - 1) = StrParaDbl(DispMaqQua.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137331

    Saida_Celula_DispMaqQua = SUCESSO

    Exit Function

Erro_Saida_Celula_DispMaqQua:

    Saida_Celula_DispMaqQua = gErr

    Select Case gErr

        Case 137329 To 137331
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144380)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispCTQui(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispCTQui

    Set objGridInt.objControle = DispCTQui

    'Se o campo foi preenchido
    If Len(DispCTQui.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispCTQui.Text, GridTurnosDias.Row, iGrid_DispCTQui_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137332
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137333

    Saida_Celula_DispCTQui = SUCESSO

    Exit Function

Erro_Saida_Celula_DispCTQui:

    Saida_Celula_DispCTQui = gErr

    Select Case gErr

        Case 137332, 137333
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144381)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispMaqQui(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispMaqQui

    Set objGridInt.objControle = DispMaqQui

    'Se o campo foi preenchido
    If Len(DispMaqQui.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispMaqQui.Text, GridDisponibilidadeMaquina.Row, iGrid_DispMaqQui_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137334
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridDisponibilidadeMaquina.Row - GridDisponibilidadeMaquina.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    gcolMaquinas.Item(GridItens.Row).colTurnos.Item(GridDisponibilidadeMaquina.Row).dQtdHoras(iGrid_DispMaqQui_Col - 1) = StrParaDbl(DispMaqQui.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137336

    Saida_Celula_DispMaqQui = SUCESSO

    Exit Function

Erro_Saida_Celula_DispMaqQui:

    Saida_Celula_DispMaqQui = gErr

    Select Case gErr

        Case 137334 To 137336
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144382)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispCTSex(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispCTSex

    Set objGridInt.objControle = DispCTSex

    'Se o campo foi preenchido
    If Len(DispCTSex.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispCTSex.Text, GridTurnosDias.Row, iGrid_DispCTSex_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137337
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137338

    Saida_Celula_DispCTSex = SUCESSO

    Exit Function

Erro_Saida_Celula_DispCTSex:

    Saida_Celula_DispCTSex = gErr

    Select Case gErr

        Case 137337, 137338
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144383)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispMaqSex(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispMaqSex

    Set objGridInt.objControle = DispMaqSex

    'Se o campo foi preenchido
    If Len(DispMaqSex.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispMaqSex.Text, GridDisponibilidadeMaquina.Row, iGrid_DispMaqSex_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137339
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridDisponibilidadeMaquina.Row - GridDisponibilidadeMaquina.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    gcolMaquinas.Item(GridItens.Row).colTurnos.Item(GridDisponibilidadeMaquina.Row).dQtdHoras(iGrid_DispMaqSex_Col - 1) = StrParaDbl(DispMaqSex.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137341

    Saida_Celula_DispMaqSex = SUCESSO

    Exit Function

Erro_Saida_Celula_DispMaqSex:

    Saida_Celula_DispMaqSex = gErr

    Select Case gErr

        Case 137339 To 137341
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144384)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispCTSab(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispCTSab

    Set objGridInt.objControle = DispCTSab

    'Se o campo foi preenchido
    If Len(DispCTSab.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispCTSab.Text, GridTurnosDias.Row, iGrid_DispCTSab_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137342
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137343

    Saida_Celula_DispCTSab = SUCESSO

    Exit Function

Erro_Saida_Celula_DispCTSab:

    Saida_Celula_DispCTSab = gErr

    Select Case gErr

        Case 137342, 137343
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144385)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DispMaqSab(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DispMaqSab

    Set objGridInt.objControle = DispMaqSab

    'Se o campo foi preenchido
    If Len(DispMaqSab.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(DispMaqSab.Text, GridDisponibilidadeMaquina.Row, iGrid_DispMaqSab_Col, objGridInt)
        If lErro <> SUCESSO Then gError 137344
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridDisponibilidadeMaquina.Row - GridDisponibilidadeMaquina.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    gcolMaquinas.Item(GridItens.Row).colTurnos.Item(GridDisponibilidadeMaquina.Row).dQtdHoras(iGrid_DispMaqSab_Col - 1) = StrParaDbl(DispMaqSab.Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137346

    Saida_Celula_DispMaqSab = SUCESSO

    Exit Function

Erro_Saida_Celula_DispMaqSab:

    Saida_Celula_DispMaqSab = gErr

    Select Case gErr

        Case 137344 To 137346
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144386)

    End Select

    Exit Function

End Function

Function Horas_Turno_Critica(sDispHorasTurno As String, iGridLinha As Integer, iGridColuna As Integer, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dQtdeTotalHoras As Double

On Error GoTo Erro_Horas_Turno_Critica

    'Critica a Quantidade de Horas
    lErro = Valor_Positivo_Critica(sDispHorasTurno)
    If lErro <> SUCESSO Then gError 137347
    
    'Efetua a Somatória da Quantidade Total de Horas
    For iIndice = 1 To objGridInt.iLinhasExistentes
    
        'se é a linha que estou alterando ...
        If iIndice = iGridLinha Then
        
            'despreza a hora do grid e acumula a que esta sendo alterada
            dQtdeTotalHoras = dQtdeTotalHoras + StrParaDbl(sDispHorasTurno)
        
        Else
            
            'escolhe o grid que está sendo mexido
            If objGridInt.objGrid.Name = GridTurnosDias.Name Then
            
                'acumula as horas
                dQtdeTotalHoras = dQtdeTotalHoras + StrParaDbl(GridTurnosDias.TextMatrix(iIndice, iGridColuna))
            
            ElseIf objGridInt.objGrid.Name = GridDisponibilidadeMaquina.Name Then
            
                'acumula as horas
                dQtdeTotalHoras = dQtdeTotalHoras + StrParaDbl(GridDisponibilidadeMaquina.TextMatrix(iIndice, iGridColuna))
            
            End If
        
        End If
    
    Next
    
    'Verifica se a Somatória das Horas é maior que 24 horas
    If dQtdeTotalHoras > HORAS_DO_DIA Then gError 137348
    
    Horas_Turno_Critica = SUCESSO
    
    Exit Function
    
Erro_Horas_Turno_Critica:

    Horas_Turno_Critica = gErr
    
    Select Case gErr
    
        Case 137347
        
        Case 137348
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDEHORASGRID_EXCEDE_DIA", gErr, dQtdeTotalHoras, DiaDaSemana(iGridColuna))
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144387)
    
    End Select
    
    Exit Function

End Function

Function DiaDaSemana(iDia As Integer) As String

    Select Case iDia
    
        Case Is = 1
            DiaDaSemana = Chr(34) & "Domingo" & Chr(34)
        
        Case Is = 2
            DiaDaSemana = Chr(34) & "Segunda-feira" & Chr(34)
        
        Case Is = 3
            DiaDaSemana = Chr(34) & "Terça-feira" & Chr(34)
        
        Case Is = 4
            DiaDaSemana = Chr(34) & "Quarta-feira" & Chr(34)
        
        Case Is = 5
            DiaDaSemana = Chr(34) & "Quinta-feira" & Chr(34)
        
        Case Is = 6
            DiaDaSemana = Chr(34) & "Sexta-feira" & Chr(34)
        
        Case Is = 7
            DiaDaSemana = Chr(34) & "Sábado" & Chr(34)
    
    End Select

End Function

Function Trata_DiasUteisNoGrid(iColuna As Integer) As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Trata_DiasUteisNoGrid

    'Para cada linha do grid ...
    For iLinha = 0 To objGridTurnosDias.iLinhasExistentes
    
        'Se é Dia Útil
        If Dias.Selected(iColuna) = True Then
        
            'e Quantidade de Horas por Turno está Preenchida ...
            If Len(HorasTurno.Text) <> 0 Then
                
                'Preenche com a quantidade default
                GridTurnosDias.TextMatrix(iLinha + 1, iColuna + 1) = Formata_Estoque(StrParaDbl(HorasTurno.Text))
            
            End If
        
        Else
        
            'senão é Dia Útil... Limpa a célula do grid
            GridTurnosDias.TextMatrix(iLinha + 1, iColuna + 1) = ""
        
        End If
    
    Next
    
    Trata_DiasUteisNoGrid = SUCESSO
    
    Exit Function
    
Erro_Trata_DiasUteisNoGrid:

    Trata_DiasUteisNoGrid = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144388)
    
    End Select

    Exit Function
    
End Function

Function Marca_Default_Dias() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Marca_Default_Dias

    For iIndice = SEGUNDA To SEXTA

        Dias.Selected(iIndice - 1) = True
    
    Next
    
    Dias.Selected(SABADO - 1) = False
    Dias.Selected(DOMINGO - 1) = False

    Marca_Default_Dias = SUCESSO
    
    Exit Function
    
Erro_Marca_Default_Dias:

    Marca_Default_Dias = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144389)
    
    End Select

    Exit Function
    
End Function

Function Preenche_GridTurnosDias_Padrao() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim iDias As Integer

On Error GoTo Erro_Preenche_GridTurnosDias_Padrao

    'Limpa o Grid ...
    Call Grid_Limpa(objGridTurnosDias)
    
    'Para cada Turno...
    For iLinha = 1 To StrParaInt(QtdeTurnos.Text)
    
        'e para cada dia da Semana
        For iDias = DOMINGO To SABADO
        
            'se for Dia Útil...
            If Dias.Selected(iDias - 1) = True Then
            
                'Preenche com a quantidade default
                GridTurnosDias.TextMatrix(iLinha, iDias) = Formata_Estoque(StrParaDbl(HorasTurno.Text))
            
            End If
        
        Next
    
    Next
    
    'atualiza as linhas do grid
    objGridTurnosDias.iLinhasExistentes = StrParaInt(QtdeTurnos.Text) - 1
    
    Preenche_GridTurnosDias_Padrao = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridTurnosDias_Padrao:

    Preenche_GridTurnosDias_Padrao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144390)
    
    End Select

    Exit Function

End Function

Function Preenche_GridDispMaquina() As Long

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim objCTMaquinas As New ClassCTMaquinas
Dim colTurnos As New Collection
Dim bMaquinaExistente As Boolean

On Error GoTo Erro_Preenche_GridDispMaquina

    'Limpa o Grid ...
    Call Grid_Limpa(objGridDisponibilidadeMaquina)
    
    'se linha do Grid de Maquinas esta preenchida...
    If GridItens.Row > 0 And Len(GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)) <> 0 Then
    
        CodigoItem.Text = GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)
    
        Set objMaquinas = New ClassMaquinas
    
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", CodigoItem, objMaquinas)
        If lErro <> SUCESSO Then gError 137349
    
        'para cada objeto da coleção...
        For Each objCTMaquinas In gcolMaquinas
        
            If objCTMaquinas.lNumIntDocMaq = objMaquinas.lNumIntDoc Then
                
                'pega a coleção de turnos do objeto
                Set colTurnos = objCTMaquinas.colTurnos
                bMaquinaExistente = True
                Exit For
                
            End If
        
        Next
        
        'se a máquina já existe na coleção ...
        If bMaquinaExistente Then
            
            'preenche o grid com os dados da coleção global
            lErro = Preenche_GridDispMaquina_Existente(colTurnos)
            If lErro <> SUCESSO Then gError 137350
            
        Else
        
            'preenche o grid com os dados do GridTurnosDias
            lErro = Preenche_GridDispMaquina_Padrao(colTurnos)
            If lErro <> SUCESSO Then gError 137351
            
            'Prepara o objeto para por na coleção
            Set objCTMaquinas = New ClassCTMaquinas
            
            objCTMaquinas.lNumIntDocMaq = objMaquinas.lNumIntDoc
            
            'seta a coleção do objeto para a coleção posta no grid
            Set objCTMaquinas.colTurnos = colTurnos
            
            'adiciona a coleção de máquinas (global)
            gcolMaquinas.Add objCTMaquinas, "X" & Right((1000 + GridItens.Row), 3)
        
        End If
        
        FrameDispMaquina.Caption = "Disponibilidade em Horas (" & objMaquinas.sNomeReduzido & ")"
    
    Else
    
        FrameDispMaquina.Caption = "Disponibilidade em Horas"
    
    End If
    
    Preenche_GridDispMaquina = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridDispMaquina:

    Preenche_GridDispMaquina = gErr
    
    Select Case gErr
    
        Case 137349 To 137351
            'erros tratados nas rotinas chamadas
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144391)
    
    End Select
    
    Exit Function

End Function

Function Preenche_GridDispMaquina_Padrao(colTurnos As Collection) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim iDias As Integer
Dim objTurno As ClassTurno
Dim dHorasGridDias As Double

On Error GoTo Erro_Preenche_GridDispMaquina_Padrao

    'Para cada Turno...
    For iLinha = 1 To StrParaInt(QtdeTurnos.Text)
    
        'prepara o objeto para por na coleção
        Set objTurno = New ClassTurno
        
        objTurno.iTurno = iLinha
        GridDisponibilidadeMaquina.TextMatrix(iLinha, iGrid_TurnoMaq_Col) = iLinha
    
        'e para cada dia da Semana
        For iDias = DOMINGO To SABADO
        
            'se for Dia Útil...
            If Dias.Selected(iDias - 1) = True Then
            
                If Len(GridTurnosDias.TextMatrix(iLinha, iDias)) <> 0 Then
                
                    dHorasGridDias = StrParaDbl(GridTurnosDias.TextMatrix(iLinha, iDias))
            
                    'Preenche com a quantidade default do GridTurnoDias
                    GridDisponibilidadeMaquina.TextMatrix(iLinha, iDias + 1) = Formata_Estoque(dHorasGridDias)
                
                    objTurno.dQtdHoras(iDias) = dHorasGridDias
                    
                End If
            
            End If
        
        Next
        
        'adiciona o objeto a coleção
        colTurnos.Add objTurno
    
    Next
    
    'atualiza as linhas do grid
    objGridDisponibilidadeMaquina.iLinhasExistentes = StrParaInt(QtdeTurnos.Text)
    
    Preenche_GridDispMaquina_Padrao = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridDispMaquina_Padrao:

    Preenche_GridDispMaquina_Padrao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144392)
    
    End Select
    
    Exit Function

End Function

Function Preenche_GridDispMaquina_Existente(colTurnos As Collection) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim iDias As Integer
Dim objTurno As New ClassTurno
Dim dHorasGridDias As Double

On Error GoTo Erro_Preenche_GridDispMaquina_Existente

    iLinha = 1

    'Para cada Turno...
    For Each objTurno In colTurnos
    
        GridDisponibilidadeMaquina.TextMatrix(iLinha, iGrid_TurnoMaq_Col) = objTurno.iTurno
    
        'e para cada dia da Semana
        For iDias = DOMINGO To SABADO
            
            If objTurno.dQtdHoras(iDias) <> 0 Then
            
                'Preenche com a quantidade que esta na coleção
                GridDisponibilidadeMaquina.TextMatrix(iLinha, iDias + 1) = Formata_Estoque(objTurno.dQtdHoras(iDias))
            
            End If
        
        Next
        
        iLinha = iLinha + 1
        
    Next
    
    'atualiza as linhas do grid
    objGridDisponibilidadeMaquina.iLinhasExistentes = colTurnos.Count
    
    Preenche_GridDispMaquina_Existente = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridDispMaquina_Existente:

    Preenche_GridDispMaquina_Existente = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144393)
    
    End Select
    
    Exit Function

End Function

'###############################################

Private Function Inicializa_GridOperadores(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Tipo M.O.")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Quantidade")

    'Controles que participam do Grid
    objGrid.colCampo.Add (CodTipoMO.Name)
    objGrid.colCampo.Add (DescricaoMO.Name)
    objGrid.colCampo.Add (QuantidadeMO.Name)

    'Colunas do Grid
    iGrid_CodigoTipoMO_Col = 1
    iGrid_DescricaoMo_Col = 2
    iGrid_QuantidadeMO_Col = 3

    objGrid.objGrid = GridOperadores

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 11

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridOperadores = SUCESSO

End Function

Private Sub GridOperadores_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridOperadores, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridOperadores, iAlterado)
        End If

End Sub

Private Sub GridOperadores_GotFocus()
    
    Call Grid_Recebe_Foco(objGridOperadores)

End Sub

Private Sub GridOperadores_EnterCell()

    Call Grid_Entrada_Celula(objGridOperadores, iAlterado)

End Sub

Private Sub GridOperadores_LeaveCell()
    
    Call Saida_Celula(objGridOperadores)

End Sub

Private Sub GridOperadores_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridOperadores, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOperadores, iAlterado)
    End If

End Sub

Private Sub GridOperadores_RowColChange()

    Call Grid_RowColChange(objGridOperadores)

End Sub

Private Sub GridOperadores_Scroll()

    Call Grid_Scroll(objGridOperadores)

End Sub

Private Sub CodTipoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodTipoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperadores)

End Sub

Private Sub CodTipoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperadores)

End Sub

Private Sub CodTipoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperadores.objControle = CodTipoMO
    lErro = Grid_Campo_Libera_Foco(objGridOperadores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperadores)

End Sub

Private Sub DescricaoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperadores)

End Sub

Private Sub DescricaoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperadores.objControle = DescricaoMO
    lErro = Grid_Campo_Libera_Foco(objGridOperadores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperadores)

End Sub

Private Sub QuantidadeMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperadores)

End Sub

Private Sub QuantidadeMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperadores.objControle = QuantidadeMO
    lErro = Grid_Campo_Libera_Foco(objGridOperadores)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_QuantidadeMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula QuantidadeMO do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantidadeMO

    Set objGridInt.objControle = QuantidadeMO

    'Se o campo foi preenchido
    If Len(Trim(QuantidadeMO.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(QuantidadeMO.Text)
        If lErro <> SUCESSO Then gError 139078
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperadores.Row - GridOperadores.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139079

    Saida_Celula_QuantidadeMO = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeMO:

    Saida_Celula_QuantidadeMO = gErr

    Select Case gErr
        
        Case 139078, 139079
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144394)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CodTipoMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CodTipoMO do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim sCodTipoMO As String
Dim objTipoMO As ClassTiposDeMaodeObra

On Error GoTo Erro_Saida_Celula_CodTipoMO

    Set objGridInt.objControle = CodTipoMO

    'Se o campo foi preenchido
    If Len(Trim(CodTipoMO.Text)) > 0 Then
    
        'Critica a Codigo
        lErro = Inteiro_Critica(CodTipoMO.Text)
        If lErro <> SUCESSO Then gError 139080
        
        'Verifica se há algum Tipo de MO repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridOperadores.Row Then
                                                    
                If GridOperadores.TextMatrix(iLinha, iGrid_CodigoTipoMO_Col) = CodTipoMO.Text Then
                    sCodTipoMO = CodTipoMO.Text
                    CodTipoMO.Text = ""
                    gError 139081
                End If
                    
            End If
                           
        Next
        
        Set objTipoMO = New ClassTiposDeMaodeObra
        
        objTipoMO.iCodigo = StrParaInt(CodTipoMO.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTipoMO)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 139082
        
        If lErro <> SUCESSO Then gError 141953
        
        GridOperadores.TextMatrix(GridOperadores.Row, iGrid_DescricaoMo_Col) = objTipoMO.sDescricao
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridOperadores.Row - GridOperadores.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 139083

    Saida_Celula_CodTipoMO = SUCESSO

    Exit Function

Erro_Saida_Celula_CodTipoMO:

    Saida_Celula_CodTipoMO = gErr

    Select Case gErr
        
        Case 139081
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_REPETIDO", gErr, sCodTipoMO, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 139080, 139082, 139083
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 141953
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTipoMO.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144395)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

'#################################################################
'Inserido por Wagner 18/11/05
Private Sub BotaoParadas_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas
Dim objCTMaquinasParadas As ClassCTMaquinasParadas

On Error GoTo Erro_BotaoParadas_Click

    'Verifica se existe uma linha do grid selecionada
    If GridItens.Row = 0 Then gError 140856
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)) = 0 Then gError 140857
    
    'Verifica se o Código do CT está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 140858
    
    objCentrodeTrabalho.lCodigo = StrParaLong(Codigo.Text)
    objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
            
    'Lê o CentrodeTrabalho que está sendo Passado
    lErro = CF("CentrodeTrabalho_Le", objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 134449 Then gError 140859
    
    'se CT não está cadastrado -> Erro
    If lErro <> SUCESSO Then gError 140860
            
     'Inicializa o controle para ler a máquina
    CodigoItem.Text = GridItens.TextMatrix(GridItens.Row, iGrid_CodigoItem_Col)

    Set objMaquinas = New ClassMaquinas

    'Verifica a existencia da máquina e lê seu NumIntDoc
    lErro = CF("TP_Maquina_Le", CodigoItem, objMaquinas)
    If lErro <> SUCESSO Then gError 140861
    
    'Inicializa o obj da tela a ser chamada
    Set objCTMaquinasParadas = New ClassCTMaquinasParadas
    
    objCTMaquinasParadas.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    objCTMaquinasParadas.lNumIntDocMaq = objMaquinas.lNumIntDoc
    
    'Chama a tela ...
    lErro = Chama_Tela("CTMaquinasParadas", objCTMaquinasParadas)
    If lErro <> SUCESSO Then gError 140862
    
    Exit Sub
    
Erro_BotaoParadas_Click:

    Select Case gErr
    
        Case 140856
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRIDMAQUINA_NAO_SELECIONADA", gErr)
        
        Case 140857
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 140858
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CENTRODETRABALHO_NAO_PREENCHIDO", gErr)
        
        Case 140859, 140861, 140862
            'erro tratado nas rotinas chamadas
        
        Case 140860
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_CADASTRADO", gErr, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144396)
    
    End Select

    Exit Sub

End Sub
'##########################################################################

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl AvisoPagtoCPPorEmailOcx 
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LockControls    =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5610
      Index           =   1
      Left            =   135
      TabIndex        =   23
      Top             =   750
      Width           =   9165
      Begin VB.ComboBox Modelo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "AvisoPagtoCPPorEmail.ctx":0000
         Left            =   2205
         List            =   "AvisoPagtoCPPorEmail.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   5880
      End
      Begin VB.Frame Frame3 
         Caption         =   "Filtros"
         Height          =   4560
         Left            =   705
         TabIndex        =   27
         Top             =   870
         Width           =   7815
         Begin VB.CheckBox IgnoraJaEnviados 
            Caption         =   "Ignorar parcelas cujo email j� tenha sido enviado"
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
            Left            =   240
            TabIndex        =   2
            Top             =   690
            Width           =   5205
         End
         Begin VB.CheckBox EmailValido 
            Caption         =   "S� trazer dados de fornecedores que possuam email v�lido"
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
            Left            =   240
            TabIndex        =   1
            Top             =   330
            Width           =   5700
         End
         Begin VB.Frame Frame6 
            Caption         =   "N� do T�tulo"
            Height          =   1575
            Left            =   2955
            TabIndex        =   48
            Top             =   1185
            Width           =   2175
            Begin MSMask.MaskEdBox TituloInic 
               Height          =   300
               Left            =   585
               TabIndex        =   3
               Top             =   420
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "999999999"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TituloFim 
               Height          =   300
               Left            =   585
               TabIndex        =   4
               Top             =   960
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "999999999"
               PromptChar      =   " "
            End
            Begin VB.Label Label21 
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
               Height          =   255
               Left            =   225
               TabIndex        =   50
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label22 
               Caption         =   "At�:"
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
               Left            =   180
               TabIndex        =   49
               Top             =   990
               Width           =   375
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Tipo de Documento"
            Height          =   1425
            Left            =   2955
            TabIndex        =   47
            Top             =   2970
            Width           =   4575
            Begin VB.ComboBox TipoDocSeleciona 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "AvisoPagtoCPPorEmail.ctx":003F
               Left            =   1410
               List            =   "AvisoPagtoCPPorEmail.ctx":0041
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   795
               Width           =   2820
            End
            Begin VB.OptionButton TipoDocTodos 
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
               Left            =   315
               TabIndex        =   7
               Top             =   360
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.OptionButton TipoDocApenas 
               Caption         =   "Apenas:"
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
               Left            =   300
               TabIndex        =   8
               Top             =   840
               Width           =   1050
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data da Baixa"
            Height          =   1575
            Left            =   270
            TabIndex        =   31
            Top             =   1185
            Width           =   2475
            Begin MSComCtl2.UpDown UpDownBaixaDe 
               Height          =   300
               Left            =   1590
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   435
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataBaixaDe 
               Height          =   300
               Left            =   525
               TabIndex        =   59
               Top             =   435
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownBaixaAte 
               Height          =   300
               Left            =   1590
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   945
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataBaixaAte 
               Height          =   300
               Left            =   525
               TabIndex        =   61
               Top             =   945
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "At�:"
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
               Index           =   2
               Left            =   120
               TabIndex        =   33
               Top             =   1005
               Width           =   360
            End
            Begin VB.Label Label1 
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
               Index           =   0
               Left            =   150
               TabIndex        =   32
               Top             =   435
               Width           =   315
            End
         End
         Begin VB.Frame FrameCliente 
            Caption         =   "Fornecedor"
            Height          =   1425
            Left            =   270
            TabIndex        =   28
            Top             =   2970
            Width           =   2475
            Begin MSMask.MaskEdBox ClienteInicial 
               Height          =   315
               Left            =   525
               TabIndex        =   5
               Top             =   345
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   556
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ClienteFinal 
               Height          =   315
               Left            =   525
               TabIndex        =   6
               Top             =   840
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   556
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label LabelClienteDe 
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
               Height          =   255
               Left            =   150
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   30
               Top             =   390
               Width           =   360
            End
            Begin VB.Label LabelClienteAte 
               Caption         =   "At�:"
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
               Height          =   255
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   29
               Top             =   885
               Width           =   435
            End
         End
      End
      Begin VB.Label LabelModelo 
         Caption         =   "Modelo de email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   690
         TabIndex        =   62
         Top             =   390
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5625
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   765
      Visible         =   0   'False
      Width           =   9225
      Begin MSMask.MaskEdBox AnexoGrid 
         Height          =   255
         Left            =   180
         TabIndex        =   57
         Top             =   120
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Frame Frame7 
         Caption         =   "Email"
         Height          =   2460
         Left            =   1905
         TabIndex        =   52
         Top             =   3120
         Width           =   7275
         Begin VB.TextBox Anexo 
            Height          =   285
            Left            =   1260
            MaxLength       =   250
            TabIndex        =   16
            Top             =   1185
            Width           =   5835
         End
         Begin VB.TextBox Email 
            Height          =   285
            Left            =   1260
            MaxLength       =   8000
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   165
            Width           =   5835
         End
         Begin VB.TextBox Cc 
            Height          =   285
            Left            =   1260
            MaxLength       =   8000
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   495
            Width           =   5835
         End
         Begin VB.TextBox Assunto 
            Height          =   285
            Left            =   1260
            MaxLength       =   250
            TabIndex        =   15
            Top             =   825
            Width           =   5835
         End
         Begin VB.TextBox Mensagem 
            Height          =   870
            Left            =   1260
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   1530
            Width           =   5835
         End
         Begin VB.Label Label2 
            Caption         =   "Anexo:"
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
            Index           =   2
            Left            =   555
            TabIndex        =   56
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label4 
            Caption         =   "Para:"
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
            Left            =   705
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   51
            Top             =   195
            Width           =   480
         End
         Begin VB.Label LabelCc 
            Caption         =   "Cc:"
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
            Left            =   855
            TabIndex        =   55
            Top             =   525
            Width           =   330
         End
         Begin VB.Label Label2 
            Caption         =   "Assunto:"
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
            Index           =   1
            Left            =   420
            TabIndex        =   54
            Top             =   855
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Mensagem:"
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
            Index           =   1
            Left            =   210
            TabIndex        =   53
            Top             =   1515
            Width           =   1020
         End
      End
      Begin MSMask.MaskEdBox MensagemGrid 
         Height          =   255
         Left            =   3360
         TabIndex        =   46
         Top             =   2235
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AssuntoGrid 
         Height          =   255
         Left            =   3660
         TabIndex        =   45
         Top             =   2550
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CCGrid 
         Height          =   255
         Left            =   2985
         TabIndex        =   44
         Top             =   1380
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox EmailGrid 
         Height          =   255
         Left            =   2490
         TabIndex        =   43
         Top             =   1950
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.ComboBox Carta 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "AvisoPagtoCPPorEmail.ctx":0043
         Left            =   3090
         List            =   "AvisoPagtoCPPorEmail.ctx":004D
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   420
         Width           =   2235
      End
      Begin MSMask.MaskEdBox Atraso 
         Height          =   255
         Left            =   7170
         TabIndex        =   41
         Top             =   2355
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Saldo 
         Height          =   225
         Left            =   7125
         TabIndex        =   40
         Top             =   2130
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   225
         Left            =   7185
         TabIndex        =   39
         Top             =   1860
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   225
         Left            =   7140
         TabIndex        =   38
         Top             =   1605
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Filial 
         Height          =   255
         Left            =   7050
         TabIndex        =   37
         Top             =   1320
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   255
         Left            =   4635
         TabIndex        =   36
         Top             =   1020
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Tipo 
         Height          =   255
         Left            =   7185
         TabIndex        =   35
         Top             =   675
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   255
         Left            =   7170
         TabIndex        =   34
         Top             =   390
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CheckBox Selecionado 
         DragMode        =   1  'Automatic
         Height          =   270
         Left            =   1275
         TabIndex        =   26
         Top             =   1710
         Width           =   555
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   255
         Left            =   7155
         TabIndex        =   25
         Top             =   75
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   60
         Picture         =   "AvisoPagtoCPPorEmail.ctx":0082
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3450
         Width           =   1725
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   45
         Picture         =   "AvisoPagtoCPPorEmail.ctx":109C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4200
         Width           =   1725
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   3090
         Left            =   45
         TabIndex        =   10
         Top             =   0
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   5450
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7635
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   45
      Width           =   1755
      Begin VB.CommandButton BotaoEmail 
         Height          =   360
         Left            =   120
         Picture         =   "AvisoPagtoCPPorEmail.ctx":227E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Enviar email"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   645
         Picture         =   "AvisoPagtoCPPorEmail.ctx":2C20
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1155
         Picture         =   "AvisoPagtoCPPorEmail.ctx":3152
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6090
      Left            =   60
      TabIndex        =   21
      Top             =   360
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10742
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sele��o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Baixas"
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
Attribute VB_Name = "AvisoPagtoCPPorEmailOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTCobrancaPorEmail
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTCobrancaPorEmail
    Set objCT.objUserControl = Me
End Sub

'Private Sub AtrasoAte_Change()
'     Call objCT.AtrasoAte_Change
'End Sub
'
'Private Sub AtrasoDe_Change()
'     Call objCT.AtrasoDe_Change
'End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
     Call objCT.TabStrip1_BeforeClick(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros(TIPOTELA_EMAIL_AVISO_PAGTO_CP)
End Function

Private Sub GridItens_Click()
     Call objCT.GridItens_Click
End Sub

Private Sub GridItens_GotFocus()
     Call objCT.GridItens_GotFocus
End Sub

Private Sub GridItens_EnterCell()
     Call objCT.GridItens_EnterCell
End Sub

Private Sub GridItens_LeaveCell()
     Call objCT.GridItens_LeaveCell
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)
     Call objCT.GridItens_KeyPress(KeyAscii)
End Sub

Private Sub GridItens_RowColChange()
     Call objCT.GridItens_RowColChange
End Sub

Private Sub GridItens_Scroll()
     Call objCT.GridItens_Scroll
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_GotFocus()
     Call objCT.Cliente_GotFocus
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
     Call objCT.Cliente_KeyPress(KeyAscii)
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub Filial_GotFocus()
     Call objCT.Filial_GotFocus
End Sub

Private Sub Filial_KeyPress(KeyAscii As Integer)
     Call objCT.Filial_KeyPress(KeyAscii)
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub Selecionado_Change()
     Call objCT.Selecionado_Change
End Sub

Private Sub Selecionado_GotFocus()
     Call objCT.Selecionado_GotFocus
End Sub

Private Sub Selecionado_KeyPress(KeyAscii As Integer)
     Call objCT.Selecionado_KeyPress(KeyAscii)
End Sub

Private Sub Selecionado_Validate(Cancel As Boolean)
     Call objCT.Selecionado_Validate(Cancel)
End Sub

Private Sub Numero_Change()
     Call objCT.Numero_Change
End Sub

Private Sub Numero_GotFocus()
     Call objCT.Numero_GotFocus
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
     Call objCT.Numero_KeyPress(KeyAscii)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
     Call objCT.Numero_Validate(Cancel)
End Sub

Private Sub Parcela_Change()
     Call objCT.Parcela_Change
End Sub

Private Sub Parcela_GotFocus()
     Call objCT.Parcela_GotFocus
End Sub

Private Sub Parcela_KeyPress(KeyAscii As Integer)
     Call objCT.Parcela_KeyPress(KeyAscii)
End Sub

Private Sub Parcela_Validate(Cancel As Boolean)
     Call objCT.Parcela_Validate(Cancel)
End Sub

Private Sub Tipo_Change()
     Call objCT.Tipo_Change
End Sub

Private Sub Tipo_GotFocus()
     Call objCT.Tipo_GotFocus
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)
     Call objCT.Tipo_KeyPress(KeyAscii)
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub DataVencimento_Change()
     Call objCT.DataVencimento_Change
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimento_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
End Sub

Private Sub Valor_Change()
     Call objCT.Valor_Change
End Sub

Private Sub Valor_GotFocus()
     Call objCT.Valor_GotFocus
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Valor_KeyPress(KeyAscii)
End Sub

Private Sub Valor_Validate(Cancel As Boolean)
     Call objCT.Valor_Validate(Cancel)
End Sub

Private Sub Saldo_Change()
     Call objCT.Saldo_Change
End Sub

Private Sub Saldo_GotFocus()
     Call objCT.Saldo_GotFocus
End Sub

Private Sub Saldo_KeyPress(KeyAscii As Integer)
     Call objCT.Saldo_KeyPress(KeyAscii)
End Sub

Private Sub Saldo_Validate(Cancel As Boolean)
     Call objCT.Saldo_Validate(Cancel)
End Sub

Private Sub Atraso_Change()
     Call objCT.Atraso_Change
End Sub

Private Sub Atraso_GotFocus()
     Call objCT.Atraso_GotFocus
End Sub

Private Sub Atraso_KeyPress(KeyAscii As Integer)
     Call objCT.Atraso_KeyPress(KeyAscii)
End Sub

Private Sub Atraso_Validate(Cancel As Boolean)
     Call objCT.Atraso_Validate(Cancel)
End Sub

Private Sub Carta_Change()
     Call objCT.Carta_Change
End Sub

Private Sub Carta_GotFocus()
     Call objCT.Carta_GotFocus
End Sub

Private Sub Carta_KeyPress(KeyAscii As Integer)
     Call objCT.Carta_KeyPress(KeyAscii)
End Sub

Private Sub Carta_Validate(Cancel As Boolean)
     Call objCT.Carta_Validate(Cancel)
End Sub

Private Sub EmailGrid_Change()
     Call objCT.EmailGrid_Change
End Sub

Private Sub EmailGrid_GotFocus()
     Call objCT.EmailGrid_GotFocus
End Sub

Private Sub EmailGrid_KeyPress(KeyAscii As Integer)
     Call objCT.EmailGrid_KeyPress(KeyAscii)
End Sub

Private Sub EmailGrid_Validate(Cancel As Boolean)
     Call objCT.EmailGrid_Validate(Cancel)
End Sub

Private Sub CCGrid_Change()
     Call objCT.CCGrid_Change
End Sub

Private Sub CCGrid_GotFocus()
     Call objCT.CCGrid_GotFocus
End Sub

Private Sub CCGrid_KeyPress(KeyAscii As Integer)
     Call objCT.CCGrid_KeyPress(KeyAscii)
End Sub

Private Sub CCGrid_Validate(Cancel As Boolean)
     Call objCT.CCGrid_Validate(Cancel)
End Sub

Private Sub AssuntoGrid_Change()
     Call objCT.AssuntoGrid_Change
End Sub

Private Sub AssuntoGrid_GotFocus()
     Call objCT.AssuntoGrid_GotFocus
End Sub

Private Sub AssuntoGrid_KeyPress(KeyAscii As Integer)
     Call objCT.AssuntoGrid_KeyPress(KeyAscii)
End Sub

Private Sub AssuntoGrid_Validate(Cancel As Boolean)
     Call objCT.AssuntoGrid_Validate(Cancel)
End Sub

Private Sub AnexoGrid_Change()
     Call objCT.AnexoGrid_Change
End Sub

Private Sub AnexoGrid_GotFocus()
     Call objCT.AnexoGrid_GotFocus
End Sub

Private Sub AnexoGrid_KeyPress(KeyAscii As Integer)
     Call objCT.AnexoGrid_KeyPress(KeyAscii)
End Sub

Private Sub AnexoGrid_Validate(Cancel As Boolean)
     Call objCT.AnexoGrid_Validate(Cancel)
End Sub

Private Sub MensagemGrid_Change()
     Call objCT.MensagemGrid_Change
End Sub

Private Sub MensagemGrid_GotFocus()
     Call objCT.MensagemGrid_GotFocus
End Sub

Private Sub MensagemGrid_KeyPress(KeyAscii As Integer)
     Call objCT.MensagemGrid_KeyPress(KeyAscii)
End Sub

Private Sub MensagemGrid_Validate(Cancel As Boolean)
     Call objCT.MensagemGrid_Validate(Cancel)
End Sub

Private Sub BotaoMarcarTodos_Click()
     Call objCT.BotaoMarcarTodos_Click
End Sub

Private Sub BotaoDesmarcarTodos_Click()
     Call objCT.BotaoDesmarcarTodos_Click
End Sub

Private Sub ClienteFinal_Change()
     Call objCT.ClienteFinal_Change
End Sub

Private Sub ClienteInicial_Change()
     Call objCT.ClienteInicial_Change
End Sub

Private Sub SaldoAte_Change()
     Call objCT.SaldoAte_Change
End Sub

Private Sub SaldoDe_Change()
     Call objCT.SaldoDe_Change
End Sub

Private Sub TipoDocApenas_Click()
     Call objCT.TipoDocApenas_Click
End Sub

Private Sub TipoDocSeleciona_Change()
     Call objCT.TipoDocSeleciona_Change
End Sub

Private Sub TipoDocTodos_Click()
     Call objCT.TipoDocTodos_Click
End Sub

Private Sub TituloFim_Change()
     Call objCT.TituloFim_Change
End Sub

Private Sub TituloInic_Change()
     Call objCT.TituloInic_Change
End Sub

Private Sub BotaoEmail_Click()
     Call objCT.BotaoEmail_Click
End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)
     Call objCT.ClienteFinal_Validate(Cancel)
End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)
     Call objCT.ClienteInicial_Validate(Cancel)
End Sub

Private Sub LabelClienteAte_Click()
     Call objCT.LabelClienteAte_Click
End Sub

Private Sub LabelClienteDe_Click()
     Call objCT.LabelClienteDe_Click
End Sub

Public Function Form_Load_Ocx() As Object

    objCT.Trata_Parametros (TIPOTELA_EMAIL_AVISO_PAGTO_CP)
    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
             Set objCT.objUserControl = Nothing
             Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub DataBaixaDe_Change()
     Call objCT.DataBaixaDe_Change
End Sub

Private Sub DataBaixaDe_GotFocus()
     Call objCT.DataBaixaDe_GotFocus
End Sub

Private Sub DataBaixaDe_Validate(Cancel As Boolean)
     Call objCT.DataBaixaDe_Validate(Cancel)
End Sub

Private Sub UpDownBaixaDe_DownClick()
     Call objCT.UpDownBaixaDe_DownClick
End Sub

Private Sub UpDownBaixaDe_UpClick()
     Call objCT.UpDownBaixaDe_UpClick
End Sub

Private Sub DataBaixaAte_Change()
     Call objCT.DataBaixaAte_Change
End Sub

Private Sub DataBaixaAte_GotFocus()
     Call objCT.DataBaixaAte_GotFocus
End Sub

Private Sub DataBaixaAte_Validate(Cancel As Boolean)
     Call objCT.DataBaixaAte_Validate(Cancel)
End Sub

Private Sub UpDownBaixaAte_DownClick()
     Call objCT.UpDownBaixaAte_DownClick
End Sub

Private Sub UpDownBaixaAte_UpClick()
     Call objCT.UpDownBaixaAte_UpClick
End Sub

Private Sub EmailValido_Click()
     Call objCT.EmailValido_Click
End Sub

Private Sub IgnoraJaEnviados_Click()
     Call objCT.IgnoraJaEnviados_Click
End Sub

Private Sub Modelo_Click()
     Call objCT.IgnoraJaEnviados_Click
End Sub

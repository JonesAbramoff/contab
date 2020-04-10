VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl EmissaoBoletosOcx 
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5280
      Index           =   1
      Left            =   150
      TabIndex        =   23
      Top             =   780
      Width           =   9075
      Begin VB.Frame Frame8 
         Caption         =   "Valores"
         Height          =   1740
         Left            =   4905
         TabIndex        =   64
         Top             =   1185
         Width           =   4170
         Begin VB.CheckBox IncluirCobrBanc 
            Caption         =   "Incluir valor de cobrança bancária"
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
            Left            =   195
            TabIndex        =   13
            Top             =   735
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin MSMask.MaskEdBox ValorTafBanc 
            Height          =   315
            Left            =   2715
            TabIndex        =   14
            Top             =   1065
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor da tarifa bancária:"
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
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   65
            Top             =   1140
            Width           =   2550
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Outros"
         Height          =   1350
         Left            =   210
         TabIndex        =   63
         Top             =   3855
         Width           =   8865
         Begin VB.OptionButton EmailValido 
            Caption         =   "Ambos"
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
            Index           =   2
            Left            =   5175
            TabIndex        =   71
            Top             =   930
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton EmailValido 
            Caption         =   "Clientes sem e-mail"
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
            Index           =   1
            Left            =   3075
            TabIndex        =   70
            Top             =   930
            Width           =   2130
         End
         Begin VB.OptionButton EmailValido 
            Caption         =   "Clientes com e-mail"
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
            Index           =   0
            Left            =   960
            TabIndex        =   69
            Top             =   915
            Width           =   2130
         End
         Begin VB.CheckBox ComBoletosImpressos 
            Caption         =   "Exibe parcelas que já tiveram seus boletos impressos"
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
            Left            =   1005
            TabIndex        =   17
            Top             =   255
            Width           =   4890
         End
         Begin VB.CheckBox CheckSemValorVcto 
            Caption         =   "Deixar a data de vencimento e o valor em branco no código de barras"
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
            Left            =   1005
            TabIndex        =   18
            Top             =   615
            Width           =   6285
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data de Vencimento"
         Height          =   1035
         Left            =   4905
         TabIndex        =   31
         Top             =   0
         Width           =   4170
         Begin MSComCtl2.UpDown UpDownVencInic 
            Height          =   300
            Left            =   1635
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   570
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox VencInic 
            Height          =   300
            Left            =   570
            TabIndex        =   5
            Top             =   570
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownVencFim 
            Height          =   300
            Left            =   3780
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   570
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox VencFim 
            Height          =   300
            Left            =   2700
            TabIndex        =   7
            Top             =   570
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label17 
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
            Left            =   195
            TabIndex        =   46
            Top             =   615
            Width           =   375
         End
         Begin VB.Label Label20 
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
            Height          =   255
            Left            =   2280
            TabIndex        =   47
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Clientes"
         Height          =   780
         Left            =   210
         TabIndex        =   28
         Top             =   3060
         Width           =   8865
         Begin MSMask.MaskEdBox ClienteInicial 
            Height          =   300
            Left            =   990
            TabIndex        =   15
            Top             =   285
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteFinal 
            Height          =   300
            Left            =   5250
            TabIndex        =   16
            Top             =   300
            Width           =   3285
            _ExtentX        =   5794
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
            Left            =   4800
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   345
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
            Left            =   630
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   49
            Top             =   345
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Borderô de Cobrança"
         Height          =   1740
         Left            =   195
         TabIndex        =   24
         Top             =   1185
         Width           =   4635
         Begin VB.Frame FrameFormPre 
            Caption         =   "Formulário pré-impresso"
            Height          =   1050
            Left            =   105
            TabIndex        =   66
            Top             =   630
            Width           =   4395
            Begin VB.CheckBox GravarNossoNumero 
               Caption         =   "Gravar ""Nosso Número"" na parcela"
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
               Left            =   120
               TabIndex        =   11
               Top             =   465
               Width           =   3495
            End
            Begin VB.CheckBox FormularioPreImpresso 
               Caption         =   "Utilizando formulário pré-impresso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   3435
            End
            Begin MSMask.MaskEdBox ProxNossoNumero 
               Height          =   285
               Left            =   1455
               TabIndex        =   12
               Top             =   705
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   14
               Mask            =   "##############"
               PromptChar      =   " "
            End
            Begin VB.Label Label3 
               Caption         =   "(Sem o DV)"
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
               Left            =   3225
               TabIndex        =   68
               Top             =   720
               Width           =   1035
            End
            Begin VB.Label Label8 
               Caption         =   "A partir de:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   435
               TabIndex        =   67
               Top             =   735
               Width           =   990
            End
         End
         Begin MSMask.MaskEdBox BorderoCobrAte 
            Height          =   300
            Left            =   -10000
            TabIndex        =   25
            Top             =   360
            Visible         =   0   'False
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BorderoCobrDe 
            Height          =   300
            Left            =   990
            TabIndex        =   9
            Top             =   315
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label LabelBorderoCobrDe 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
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
            Left            =   210
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   345
            Width           =   720
         End
         Begin VB.Label LabelBorderoCobrAte 
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
            Left            =   -10000
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   51
            Top             =   405
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tipo de Cobrança"
         Height          =   1050
         Left            =   195
         TabIndex        =   60
         Top             =   -15
         Width           =   4635
         Begin VB.ComboBox ComboBanco 
            Height          =   315
            Left            =   975
            TabIndex        =   2
            Top             =   585
            Width           =   2940
         End
         Begin VB.ComboBox ComboCobrador 
            Height          =   315
            Left            =   990
            TabIndex        =   4
            Top             =   585
            Width           =   2925
         End
         Begin VB.OptionButton optCobrador 
            Caption         =   "Cobrador"
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
            Height          =   495
            Left            =   1065
            TabIndex        =   0
            Top             =   180
            Value           =   -1  'True
            Width           =   1470
         End
         Begin VB.OptionButton optBanco 
            Caption         =   "Banco"
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
            Height          =   495
            Left            =   2985
            TabIndex        =   1
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cobrador:"
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
            Left            =   60
            TabIndex        =   62
            Top             =   660
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Left            =   300
            TabIndex        =   61
            Top             =   660
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5175
      Index           =   2
      Left            =   270
      TabIndex        =   32
      Top             =   780
      Visible         =   0   'False
      Width           =   8985
      Begin VB.Frame Frame1 
         Caption         =   "Total"
         Height          =   960
         Index           =   0
         Left            =   1065
         TabIndex        =   26
         Top             =   3555
         Width           =   2250
         Begin VB.Label TotalParcelas 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   780
            TabIndex        =   52
            Top             =   585
            Width           =   1275
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Left            =   180
            TabIndex        =   53
            Top             =   615
            Width           =   510
         End
         Begin VB.Label QtdParcelas 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   780
            TabIndex        =   54
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Qtde.:"
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
            Left            =   165
            TabIndex        =   55
            Top             =   270
            Width           =   540
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Selecionados"
         Height          =   960
         Left            =   5340
         TabIndex        =   27
         Top             =   3555
         Width           =   2250
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Left            =   150
            TabIndex        =   56
            Top             =   585
            Width           =   510
         End
         Begin VB.Label TotalParcelasSelecionadas 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   765
            TabIndex        =   57
            Top             =   555
            Width           =   1275
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Qtde.:"
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
            Left            =   120
            TabIndex        =   58
            Top             =   300
            Width           =   540
         End
         Begin VB.Label QtdParcelasSelecionadas 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   765
            TabIndex        =   59
            Top             =   255
            Width           =   1275
         End
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todas"
         Height          =   585
         Left            =   3630
         Picture         =   "EmissaoBoletosOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3495
         Width           =   1440
      End
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todas"
         Height          =   585
         Left            =   3630
         Picture         =   "EmissaoBoletosOcx.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4155
         Width           =   1440
      End
      Begin VB.Frame Cobranca 
         Caption         =   "Parcelas em Aberto"
         Height          =   3315
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   8970
         Begin VB.CheckBox Selecionar 
            Height          =   225
            Left            =   105
            TabIndex        =   34
            Top             =   300
            Width           =   525
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   4920
            TabIndex        =   35
            Top             =   165
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Saldo 
            Height          =   225
            Left            =   6060
            TabIndex        =   36
            Top             =   180
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox Numero 
            Height          =   225
            Left            =   3360
            TabIndex        =   37
            Top             =   195
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Tipo 
            Height          =   225
            Left            =   2625
            TabIndex        =   38
            Top             =   210
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Parcela 
            Height          =   225
            Left            =   4140
            TabIndex        =   39
            Top             =   180
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cobrador 
            Height          =   225
            Left            =   4440
            TabIndex        =   40
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialEmpresa 
            Height          =   225
            Left            =   6480
            TabIndex        =   42
            Top             =   450
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ImpressoEm 
            Height          =   225
            Left            =   7065
            TabIndex        =   43
            Top             =   135
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Nome 
            Height          =   225
            Left            =   600
            TabIndex        =   44
            Top             =   240
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2955
            Left            =   90
            TabIndex        =   41
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   5212
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7110
      ScaleHeight     =   495
      ScaleWidth      =   2250
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   45
      Width           =   2310
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   720
         Picture         =   "EmissaoBoletosOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Executa a rotina"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimirGerar 
         Height          =   360
         Left            =   165
         Picture         =   "EmissaoBoletosOcx.ctx":263E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   1275
         Picture         =   "EmissaoBoletosOcx.ctx":2B7F
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1800
         Picture         =   "EmissaoBoletosOcx.ctx":30B1
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5775
      Left            =   105
      TabIndex        =   3
      Top             =   345
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   10186
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Emissão"
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
Attribute VB_Name = "EmissaoBoletosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'na versao light
    'esconder cobrador (será sempre a propria empresa) e bordero de/ate (nao existe bordero de cobranca)
    'incluir combo de banco que ficará escondida na versao full

'se a cobranca for centralizada em uma filial e gifilialempresa for esta filial deve trazer parcelasrec de todas as filiais,
    'senao pegar apenas as parcelasrec de titulos de gifilialempresa

'Para as parcelas selecionadas atualizar os campos IdImpressaoBoleto e DataBoleto na tabela ParcelasRec
'Id deve ser obtido de CRConfig (NUM_PROX_ID_BOLETO)

'todas as parcelas selecionadas tem que ser do mesmo cobrador

'Property Variables:
Dim m_Caption As String
Event Unload()

'Grid Parcelas:
Dim objGridParcelas As AdmGrid
Dim iGrid_Selecionar_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_DataVencimento_Col As Integer
Dim iGrid_ValorCobrado_Col As Integer
Dim iGrid_ImpressoEm_Col As Integer
Dim iGrid_Cobrador_Col As Integer
Dim iGrid_FilialEmpresa_Col As Integer

Dim iFrameAtual As Integer
Dim iFramePrincipalAlterado As Integer
Dim gcolInfoParcRec As Collection

Private Const TAB_SELECAO = 1
Private Const TAB_EMISSAO = 2

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1
Private WithEvents objEventoBorderoDe As AdmEvento
Attribute objEventoBorderoDe.VB_VarHelpID = -1
Private WithEvents objEventoBorderoAte As AdmEvento
Attribute objEventoBorderoAte.VB_VarHelpID = -1

Private Sub BorderoCobrDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(BorderoCobrDe)

End Sub


Private Sub ComboBanco_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboBanco_Click()
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EmailValido_Click(Index As Integer)
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoBorderoDe_evSelecao(obj1 As Object)

Dim objBorderoCobranca As ClassBorderoCobranca
    
    Set objBorderoCobranca = obj1
    
    BorderoCobrDe.PromptInclude = False
    
    If BorderoCobrDe.Enabled = True Then
        BorderoCobrDe.Text = objBorderoCobranca.lNumBordero
        ProxNossoNumero.Text = ""
        Call BorderoCobrDe_Validate(bSGECancelDummy)
    End If

    BorderoCobrDe.PromptInclude = True

    Me.Show

End Sub

Private Sub optBanco_Click()

    'deixa invisivel a combo banco
    ComboBanco.Visible = True
    Label2.Visible = True
    
    'deixa visivel a combo cobrador
    ComboCobrador.Visible = False
    Label1.Visible = False
    
    ComboBanco.ListIndex = -1
    ComboCobrador.ListIndex = -1
    
    Frame2.Enabled = False
    
    BorderoCobrDe.PromptInclude = False
    BorderoCobrDe.Text = ""
    BorderoCobrDe.PromptInclude = True
        
    LabelBorderoCobrDe.Enabled = False
        
    'inicializa o grid apos limpa-lo
    Call Grid_Limpa(objGridParcelas)
    
    Set objGridParcelas = New AdmGrid
    
    Call Inicializa_Grid_Parcelas(objGridParcelas)

End Sub

Private Sub optCobrador_Click()

    'deixa invisivel a combo banco
    ComboBanco.Visible = False
    Label2.Visible = False
    
    'deixa visivel a combo cobrador
    ComboCobrador.Visible = True
    Label1.Visible = True
    
    ComboCobrador.ListIndex = -1
    ComboBanco.ListIndex = -1
    
    Frame2.Enabled = True
    
    LabelBorderoCobrDe.Enabled = True
    
    BorderoCobrDe.PromptInclude = False
    BorderoCobrDe.Text = ""
    BorderoCobrDe.PromptInclude = True
    
    'inicializa o grid apos limpa-lo
    Call Grid_Limpa(objGridParcelas)
    
    Set objGridParcelas = New AdmGrid
    
    Call Inicializa_Grid_Parcelas(objGridParcelas)

End Sub

Private Sub ProxNossoNumero_GotFocus()
Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(ProxNossoNumero, iAlterado)
End Sub

Private Sub UpDownVencFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencFim_DownClick

    If Len(VencFim.ClipText) > 0 Then
        
        sData = VencFim.Text
        
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 61444

        VencFim.Text = sData

    End If
    

    Exit Sub

Erro_UpDownVencFim_DownClick:

    Select Case Err

        Case 61444

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159402)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencFim_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencFim_UpClick

    If Len(VencFim.ClipText) > 0 Then

        sData = VencFim.Text
        
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 61445

        VencFim.Text = sData
    
    End If

    Exit Sub

Erro_UpDownVencFim_UpClick:

    Select Case Err

        Case 61445

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159403)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencInic_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencInic_DownClick

    If Len(VencInic.ClipText) > 0 Then

        sData = VencInic.Text
        
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 61446

        VencInic.Text = sData
    
    End If
    

    Exit Sub

Erro_UpDownVencInic_DownClick:

    Select Case Err

        Case 61446

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159404)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencInic_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencInic_UpClick

    If Len(VencInic.ClipText) > 0 Then

        sData = VencInic.Text
        
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 61447

        VencInic.Text = sData
    
    End If

    Exit Sub

Erro_UpDownVencInic_UpClick:

    Select Case Err

        Case 61447

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159405)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

Private Sub BorderoCobrDe_Change()
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    BorderoCobrAte.Text = BorderoCobrDe.Text
    
End Sub

Private Sub BorderoCobrDe_Validate(Cancel As Boolean)

Dim lErro As Long, objCarteiraCobrador As New ClassCarteiraCobrador
Dim objBorderoCobranca As New ClassBorderoCobranca

On Error GoTo Erro_BorderoCobrDe_Validate

    'Se o Bordero de Não está Preenchido --> Sai da Sub
    If Len(Trim(BorderoCobrDe.Text)) = 0 Then Exit Sub
    
    'Verifica se o cobrador foi preenchido
    If Len(Trim(ComboCobrador.Text)) = 0 Then gError 61450
    
    objBorderoCobranca.lNumBordero = CLng(BorderoCobrDe.Text)
    objBorderoCobranca.iCobrador = Codigo_Extrai(ComboCobrador.Text)
    
    'Lê o Bordero de Cobranca
    lErro = CF("BorderoCobranca_Le_Cobrador", objBorderoCobranca)
    If lErro <> SUCESSO And lErro <> 61447 Then gError 61440
    
    'Não Encontrou o bordero
    If lErro = 61447 Then gError 61441
    
    objCarteiraCobrador.iCobrador = objBorderoCobranca.iCobrador
    objCarteiraCobrador.iCodCarteiraCobranca = objBorderoCobranca.iCodCarteiraCobranca
    lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
    If lErro <> SUCESSO And lErro <> 23551 Then gError 130030
    If lErro <> SUCESSO Then gError 130031
        
    If Len(Trim(ProxNossoNumero.Text)) = 0 Then
        ProxNossoNumero.Text = objCarteiraCobrador.sFaixaNossoNumeroProx
    End If
    
    If objCarteiraCobrador.iFormPreImp = MARCADO Then
        FrameFormPre.Enabled = True
        GravarNossoNumero.Value = vbChecked
        FormularioPreImpresso.Value = vbChecked
    Else
        FrameFormPre.Enabled = False
        ProxNossoNumero.Text = ""
        GravarNossoNumero.Value = vbUnchecked
        FormularioPreImpresso.Value = vbUnchecked
    End If
    
    Exit Sub
    
Erro_BorderoCobrDe_Validate:
        
    Cancel = True
    
    Select Case gErr
        
        Case 61357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERODE_MAIOR_BORDEROATE", gErr)
        
        Case 61440, 130030
        
        Case 130031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARTCOBR_BORDERO", gErr)
        
        Case 61441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERO_COBRANCA_NAO_CADASTRADO_COBRADOR", gErr, objBorderoCobranca.lNumBordero, objBorderoCobranca.iCobrador)
        
        Case 61450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159406)

    End Select
        
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoImprimir_Click()
    Call Imprime_Boleto
End Sub

Private Sub BotaoGerar_Click()
    Call Imprime_Boleto(True)
End Sub

Private Sub BotaoImprimirGerar_Click()
Dim lErro As Long
    lErro = Imprime_Boleto(True, True)
    If lErro = SUCESSO Then
        lErro = Imprime_Boleto(, , True)
    End If
End Sub

Private Function Imprime_Boleto(Optional bGeraArq As Boolean = False, Optional bNaoFechaTela As Boolean = False, Optional bImpressoraConfigurada As Boolean = False) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim lIdBoleto As Long, sProxNossoNumero As String
Dim objCobrador As New ClassCobrador, sBuffer As String
Dim objBanco As New ClassBanco, sArqFig1 As String, sArqFig2 As String
Dim colIDs As New Collection
Dim colNomeArq As New Collection, sBoleto As String, iIndice As Integer
Dim colInfoParcRec As Collection
Dim colInfoParcRecMarc As New Collection
Dim objInfoParcRec As ClassInfoParcRec
Dim vlIdBoleto As Variant, vsBoleto As Variant
Dim sFormatoNossoNum As String
Dim iTamNossoNumero As Integer
Dim bJaConfigImpr As Boolean

On Error GoTo Erro_Imprime_Boleto
    
    If giTipoVersao = VERSAO_FULL Then
        
        If optCobrador.Value = True Then
            'Se o cobrador está em branco --> Erro
            If Len(Trim(ComboCobrador.Text)) = 0 Then gError 61426
            'If Len(Trim(BorderoCobrDe.Text)) = 0 Then gError 130020
            
        Else
            
            If Len(Trim(ComboBanco.Text)) = 0 Then gError 111736
                    
        End If
        
''    ElseIf giTipoVersao = VERSAO_LIGHT Then
''        'Se o Banco está em branco --> Erro
''        If Len(Trim(ComboBanco.Text)) = 0 Then Error 61423
    End If
    
    'Verifica se pelo menos uma linha do Grid foi preenchida
    lErro = Verifica_Grid_Preenchido()
    If lErro <> SUCESSO Then gError 61368
    
    If optCobrador.Value = True Then
        objCobrador.iCodigo = Codigo_Extrai(ComboCobrador.Text)
        
        'Verifica qual é o Banco através do cobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 61427
        
        If lErro = 19294 Then Error 61428
        
        If (objCobrador.iCodBanco <> 0) Then
        
            objBanco.iCodBanco = objCobrador.iCodBanco
            
            'Lê o Banco para saber qual é o boleto
            lErro = CF("Banco_Le", objBanco)
            If lErro <> SUCESSO And lErro <> 16091 Then gError 61429
            
            'Se não encontrou o Banco
            If lErro = 16091 Then gError 61430
            
            lErro = CF("Banco_ObtemTamNossoNumero", objCobrador.iCodBanco, iTamNossoNumero)
            If lErro <> SUCESSO Then gError 61429
            
            sFormatoNossoNum = FormataCpoNum(0, iTamNossoNumero)
                
            If GravarNossoNumero.Value = vbChecked Then
                sProxNossoNumero = Trim(ProxNossoNumero.Text)
                If Len(sProxNossoNumero) = 0 Then gError 61430
            End If

            If Not bGeraArq Then
                'Faz a parte do BD com relação ao relatorio
                lErro = CF("TitulosRec_AtualizaBoletos_Impressao", gcolInfoParcRec, lIdBoleto, sProxNossoNumero, objBanco.iCodBanco)
                If lErro <> SUCESSO Then gError 61303
                colIDs.Add lIdBoleto
            Else
                For Each objInfoParcRec In gcolInfoParcRec
                    If objInfoParcRec.iMarcada = MARCADO Then
                        Set colInfoParcRec = New Collection
                        colInfoParcRec.Add objInfoParcRec
                        'Faz a parte do BD com relação ao relatorio
                        lErro = CF("TitulosRec_AtualizaBoletos_Impressao", colInfoParcRec, lIdBoleto, sProxNossoNumero, objBanco.iCodBanco)
                        If lErro <> SUCESSO Then gError 61303
                        colIDs.Add lIdBoleto
                        If GravarNossoNumero.Value = vbChecked Then
                            sProxNossoNumero = Format(Val(sProxNossoNumero) + 1, sFormatoNossoNum)
                        End If
                        sBoleto = gobjCRFAT.sDirBoletoGer & "BOLETO_" & Format(glEmpresa, "00") & Format(giFilialEmpresa, "00") & "_" & Format(objInfoParcRec.lNumTitulo, "000000000") & Format(objInfoParcRec.iNumParcela, "00") & "_" & Format(objInfoParcRec.dtVencimento, "YYYYMMDD") & gsExtensaoGerRelExp
                        colNomeArq.Add sBoleto
                        colInfoParcRecMarc.Add objInfoParcRec
                    End If
                Next
            End If
                   
            If FormularioPreImpresso.Value = vbUnchecked Then
            
                If Not bGeraArq Then
                    lErro = CF("ImpressaoBoletos_Prepara", gcolInfoParcRec, lIdBoleto, objCobrador, CheckSemValorVcto.Value)
                    If lErro <> SUCESSO Then gError 61429
                Else
                    iIndice = 0
                    For Each vlIdBoleto In colIDs
                        iIndice = iIndice + 1
                        lIdBoleto = vlIdBoleto
                        Set colInfoParcRec = New Collection
                        colInfoParcRec.Add colInfoParcRecMarc.Item(iIndice)
                        
                        lErro = CF("ImpressaoBoletos_Prepara", colInfoParcRec, lIdBoleto, objCobrador, CheckSemValorVcto.Value)
                        If lErro <> SUCESSO Then gError 61429
                    Next
                End If
                    
                sBuffer = String(128, 0)
                Call GetPrivateProfileString("Forprint", "DirTsks", "c:\sge\relat\", sBuffer, 128, "ADM100.INI")
                sBuffer = StringZ(sBuffer)
                If right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
                
                sArqFig1 = sBuffer & "bl" & CStr(objCobrador.iCodBanco) & "a.bmp"
                sArqFig2 = sBuffer & "bl" & CStr(objCobrador.iCodBanco) & "b.bmp"
                
            End If
            
        End If
    
    Else

        If Not bGeraArq Then
            'Faz a parte do BD com relação ao relatorio
            lErro = CF("TitulosRec_AtualizaBoletos_Impressao", gcolInfoParcRec, lIdBoleto)
            If lErro <> SUCESSO Then gError 61303
            colIDs.Add lIdBoleto
        Else
        
            For Each objInfoParcRec In gcolInfoParcRec
                If objInfoParcRec.iMarcada = MARCADO Then
                    Set colInfoParcRec = New Collection
                    colInfoParcRec.Add objInfoParcRec
                    'Faz a parte do BD com relação ao relatorio
                    lErro = CF("TitulosRec_AtualizaBoletos_Impressao", colInfoParcRec, lIdBoleto)
                    If lErro <> SUCESSO Then gError 61303
                    colIDs.Add lIdBoleto
                    sBoleto = gobjCRFAT.sDirBoletoGer & "BOLETO_" & Format(glEmpresa, "00") & Format(giFilialEmpresa, "00") & "_" & Format(objInfoParcRec.lNumTitulo, "000000000") & Format(objInfoParcRec.iNumParcela, "00") & "_" & Format(objInfoParcRec.dtVencimento, "YYYYMMDD") & gsExtensaoGerRelExp
                    colNomeArq.Add sBoleto
                    colInfoParcRecMarc.Add objInfoParcRec
                End If
            Next

        End If
    
        objBanco.iCodBanco = Codigo_Extrai(ComboBanco.Text)
        
        If (objBanco.iCodBanco <> 0) Then
    
            'Lê o Banco para saber qual é o boleto
            lErro = CF("Banco_Le", objBanco)
            If lErro <> SUCESSO And lErro <> 16091 Then gError 61431
        
            'Se não encontrou o Banco
            If lErro = 16091 Then gError 61432
        
        End If
                
    End If
    
    bJaConfigImpr = bImpressoraConfigurada
    
    iIndice = 0
    For Each vlIdBoleto In colIDs
    
        lIdBoleto = vlIdBoleto
        iIndice = iIndice + 1
        If bGeraArq Then sBoleto = colNomeArq.Item(iIndice)
        
        If Len(Trim(objBanco.sLayoutBoleto)) > 0 Then
        
            'Dispara o relatorio determinado pelo Banco
            If bJaConfigImpr Then
                objRelatorio.bConfiguraImpressora = False
            Else
                objRelatorio.bConfiguraImpressora = True
                bJaConfigImpr = True
            End If
            If bGeraArq Then
                lErro = objRelatorio.ExecutarDireto("Emissão de Boletos", "", 2, objBanco.sLayoutBoleto, "NIDBOLETO", CStr(lIdBoleto), "AARQFIG1", sArqFig1, "AARQFIG2", sArqFig2, "TTO_EMAIL", "", "TSUBJECT", "", "TALIASATTACH", "", "TMAILARQ", sBoleto)
            Else
                lErro = objRelatorio.ExecutarDireto("Emissão de Boletos", "", 0, objBanco.sLayoutBoleto, "NIDBOLETO", CStr(lIdBoleto), "AARQFIG1", sArqFig1, "AARQFIG2", sArqFig2)
            End If
            If lErro <> SUCESSO Then gError 61433
            
        Else 'Senão Imprime o default
            
            'Dispara o relatorio default
            If bJaConfigImpr Then
                objRelatorio.bConfiguraImpressora = False
            Else
                objRelatorio.bConfiguraImpressora = True
                bJaConfigImpr = True
            End If
            If bGeraArq Then
                lErro = objRelatorio.ExecutarDireto("Emissão de Boletos", "", 2, objRelatorio.sNomeTsk, "NIDBOLETO", CStr(lIdBoleto), "TTO_EMAIL", "", "TSUBJECT", "", "TALIASATTACH", "", "TMAILARQ", sBoleto)
            Else
                lErro = objRelatorio.ExecutarDireto("Emissão de Boletos", "", 0, objRelatorio.sNomeTsk, "NIDBOLETO", CStr(lIdBoleto))
            End If
            If lErro <> SUCESSO Then gError 61304
            
        End If
        
    Next
    
    Imprime_Boleto = SUCESSO
        
    If Not bNaoFechaTela Then Unload Me
        
    Exit Function
    
Erro_Imprime_Boleto:

    Imprime_Boleto = gErr

    Select Case gErr
        
        Case 61303, 61304, 61368, 61427, 61429, 61431, 61433 'Tratados nas rotinas chamadas
        
        Case 130020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_BORDERO_NAO_INFORMADO", gErr)
        
        Case 61423, 111736
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_INFORMADO", gErr)
        
        Case 61426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
        
        Case 61428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, objCobrador.iCodigo)
        
        Case 61430, 61432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", gErr, objBanco.iCodBanco)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159407)

    End Select
        
    Exit Function
    
End Function

Private Sub ClienteFinal_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Function Verifica_Grid_Preenchido() As Long
'Verifica se pelo menos uma linha do Grid foi preenchida

Dim lErro As Long
Dim objInfoParcRec As New ClassInfoParcRec
Dim iEncontrou As Integer

On Error GoTo Erro_Verifica_Grid_Preenchido
    
    iEncontrou = 0
        
    For Each objInfoParcRec In gcolInfoParcRec
            
        If objInfoParcRec.iMarcada = SELECIONAR_CHECADO Then
            iEncontrou = 1
            Exit For
        End If
    
    Next
    
    'Não encontrou Linha do Grid Selecionada
    If iEncontrou = 0 Then Error 61367
    
    Verifica_Grid_Preenchido = SUCESSO
    
    Exit Function
    
Erro_Verifica_Grid_Preenchido:

    Verifica_Grid_Preenchido = Err
    
    Select Case Err
        
        Case 61367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159408)

    End Select
    
    Exit Function
        
End Function

Private Sub ClienteFinal_Validate(Cancel As Boolean)
'Faz as criticas para o cliente

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    'Se está Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then Error 61305

        If Len(Trim(ClienteInicial.Text)) > 0 Then
            
            'Critica se o Cliente Inicial é menor que o Final
            If LCodigo_Extrai(ClienteInicial.Text) > LCodigo_Extrai(ClienteFinal.Text) Then Error 61360
        
        End If

    End If

    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case Err

        Case 61305
            ClienteFinal.SetFocus

        Case 61360
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159409)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ClienteInicial_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)
'Faz as criticas para o cliente Inicial

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then Error 61306

        If Len(Trim(ClienteFinal.Text)) > 0 Then
            
            'Critica se o cliente Inicial é menor que o Final
            If LCodigo_Extrai(ClienteInicial.Text) > LCodigo_Extrai(ClienteFinal.Text) Then Error 61359
        
        End If
        
    End If

    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case Err

        Case 61306
            ClienteInicial.SetFocus
        
        Case 61359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159410)

    End Select

End Sub

Private Sub ComboBanco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objBanco As New ClassBanco
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_ComboBanco_Validate

    'verifica se foi preenchido o ComboBanco
    If Len(Trim(ComboBanco.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox ComboBanco
    If ComboBanco.Text = ComboBanco.List(ComboBanco.ListIndex) Then Exit Sub

    'tenta Selecionar o banco com aquele codigo
    lErro = Combo_Seleciona(ComboBanco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 61424
    
    If lErro = 6730 Then
    
        objBanco.iCodBanco = iCodigo
        
        'Verifica se o banco esta no BD
        lErro = CF("Banco_Le", objBanco)
        If lErro <> SUCESSO And lErro <> 16091 Then Error 61425
        
        If lErro = 16091 Then Error 61426
        
    End If
    
    If lErro = 6731 Then Error 61427
        
    Exit Sub

Erro_ComboBanco_Validate:

    Cancel = True

    Select Case Err

        Case 61424, 61425
        
        Case 61426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", Err, ComboBanco.Text)

        Case 61427
            'Se o banco nao estiver no BD pergunta se quer criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODBANCO_INEXISTENTE", objBanco.iCodBanco)
            
            If vbMsgRes = vbYes Then
                Call Chama_Tela("Bancos", objBanco)
            End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159411)

    End Select

    Exit Sub

End Sub

Private Sub ComboCobrador_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ComboCobrador_Click()

    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComBoletosImpressos_Click()

    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelBorderoCobrDe_Click()

Dim lErro As Long
Dim objBordero As New ClassBorderoCobranca
Dim colSelecao As New Collection

On Error GoTo Erro_LabelBorderoCobrDe_Click

    'Verifica se o cobrador foi preenchido
    If Len(Trim(ComboCobrador.Text)) = 0 Then Error 61308

    If Len(Trim(BorderoCobrDe.Text)) > 0 Then objBordero.lNumBordero = CLng(BorderoCobrDe.Text)
    
    colSelecao.Add Codigo_Extrai(ComboCobrador.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("BorderoDeCobrancaLista", colSelecao, objBordero, objEventoBorderoDe)

    Exit Sub

Erro_LabelBorderoCobrDe_Click:

    Select Case Err

        Case 61308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err, Error$)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159412)

    End Select

    Exit Sub

End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente Final com o Codigo selecionado
    ClienteFinal.Text = CStr(objCliente.lCodigo)
    
    'Preenche o Cliente Final com Codigo - Descricao
    Call ClienteFinal_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente Inical com o codigo
    ClienteInicial.Text = CStr(objCliente.lCodigo)

    'Preenche o Cliente Inicial com codigo - Descricao
    Call ClienteInicial_Validate(bCancel)
    
    Me.Show
    
    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteAte_Click

    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFim)

   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159413)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteDe_Click

    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInic)

   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159414)

    End Select

    Exit Sub

End Sub

Private Sub Inicializa_Grid_Parcelas(objGridInt As AdmGrid)
'Executa a Inicialização do grid Parcelas

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Emitir")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Cobrador")
    
'    If giTipoVersao = VERSAO_FULL Then
'
'        If optCobrador.Value = True Then
'
            objGridInt.colColuna.Add ("Valor Cobrado")
'
'        Else
'
'            objGridInt.colColuna.Add ("Saldo")
'
'        End If
'
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        objGridInt.colColuna.Add ("Saldo")
'    End If
    
    objGridInt.colColuna.Add ("Impresso Em")

    'campos de edição do grid
    objGridInt.colCampo.Add (Selecionar.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Nome.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (Cobrador.Name)
    objGridInt.colCampo.Add (Saldo.Name)
    objGridInt.colCampo.Add (ImpressoEm.Name)

    iGrid_Selecionar_Col = 1
    iGrid_Tipo_Col = 2
    iGrid_Numero_Col = 3
    iGrid_Parcela_Col = 4
    iGrid_Cliente_Col = 5
    iGrid_DataVencimento_Col = 6
    iGrid_Cobrador_Col = 7
    iGrid_ValorCobrado_Col = 8
    iGrid_ImpressoEm_Col = 9
    
    If giTipoVersao = VERSAO_FULL Then
        
        If optCobrador.Value = True Then
        
            objGridInt.colColuna.Add ("Filial Empresa")
            objGridInt.colCampo.Add (FilialEmpresa.Name)
            iGrid_FilialEmpresa_Col = 10
            FilialEmpresa.Visible = True
            
        Else
        
            FilialEmpresa.Visible = False
            'FilialEmpresa.Left = -20000
        
        End If
        
    ElseIf giTipoVersao = VERSAO_LIGHT Then
        
        FilialEmpresa.Visible = False
        'FilialEmpresa.Left = -20000
    End If
    
    objGridInt.objGrid = GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 11

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10

    'largura da primeira coluna
    GridParcelas.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Sub

End Sub

Public Sub Form_UnLoad(Cancel As Integer)
        
    Set objGridParcelas = Nothing
    Set objEventoBorderoDe = Nothing
    Set objEventoBorderoAte = Nothing
    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    Set gcolInfoParcRec = Nothing
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    'seta esse cara como registro alterado
    'para que na 1a vez que ele clique no tab de parcelas
    'de erro de cobrador nao preenchido caso ele nao preencha
    'o cobrador...
    'somente pra manter a consistencia da tela
    'tulio300103
    iFramePrincipalAlterado = REGISTRO_ALTERADO

    'Inicialização
    Set objGridParcelas = New AdmGrid
    Set objEventoBorderoDe = New AdmEvento
    Set objEventoBorderoAte = New AdmEvento
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    Set gcolInfoParcRec = New Collection
    
    If giTipoVersao = VERSAO_FULL Then
        'Preenche a combo de Cobrador
        lErro = Carrega_Combo_Cobrador()
        If lErro <> SUCESSO Then gError 61309
        
        'alteracao por tulio220103
        'Preenche a combo de bancos
        lErro = Carrega_Combo_Bancos()
        If lErro <> SUCESSO Then gError 111735
        
        'deixa invisivel a combo banco
        ComboBanco.Visible = False
        Label2.Visible = False
        'fim alteracao por tulio220103
        
    ElseIf giTipoVersao = VERSAO_LIGHT Then
        'Preenche a combo de bancos
        lErro = Carrega_Combo_Bancos()
        If lErro <> SUCESSO Then gError 61420
        ComboCobrador.left = -20000
        Label1.left = -20000
        ComboCobrador.TabStop = False
        
    End If
    
    IncluirCobrBanc.Value = vbUnchecked
    ValorTafBanc.Enabled = False
    
    'Executa a Inicialização do grid Parcelas
    Call Inicializa_Grid_Parcelas(objGridParcelas)

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 61309, 61420, 111735 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159415)

    End Select

    Exit Sub

End Sub

Function Carrega_Combo_Bancos() As Long

Dim lErro As Long
Dim objCodNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome
On Error GoTo Erro_Carrega_Combo_Bancos

    'leitura dos bancos no BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_BANCO_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 61421

    'preenche listbox com nomes reduzidos dos bancos
    For Each objCodNome In colCodigoNome
        ComboBanco.AddItem objCodNome.iCodigo & SEPARADOR & objCodNome.sNome
        ComboBanco.ItemData(ComboBanco.NewIndex) = objCodNome.iCodigo
    Next

    Carrega_Combo_Bancos = SUCESSO
    
    Exit Function
    
Erro_Carrega_Combo_Bancos:

    Carrega_Combo_Bancos = Err
    
    Select Case Err
        
        Case 61421
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159416)

    End Select

    Exit Function

End Function

Private Sub ComboCobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCobrador As New ClassCobrador
Dim iCodigo As Integer

On Error GoTo Erro_ComboCobrador_Validate

    'Verifica se foi preenchida a ComboBox ComboCobrador
    If Len(Trim(ComboCobrador.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox ComboCobrador
    If ComboCobrador.Text = ComboCobrador.List(ComboCobrador.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ComboCobrador, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 61361
    
    If iCodigo = COBRADOR_PROPRIA_EMPRESA Then Error 61366
    
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCobrador.iCodigo = iCodigo

        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 61362

        If lErro <> SUCESSO Then Error 61363 'Não encontrou Cobrador no BD

        'Encontrou Cobrador no BD, coloca no Text da Combo
        ComboCobrador.Text = CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 61364

    Exit Sub

Erro_ComboCobrador_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 61361, 61362

    Case 61363  'Não encontrou Cobrador no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COBRADOR")

        If vbMsgRes = vbYes Then

            Call Chama_Tela("Cobradores", objCobrador)


        End If

    Case 61364

        lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", Err, ComboCobrador.Text)

    Case 61366
        lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_PROPRIA_EMPRESA", Err, iCodigo)
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159417)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Combo_Cobrador() As Long
'Carrega a Combo de Cobradores com todos os cobradores menos Propria Empresa

Dim lErro As Long
Dim objCobrador As ClassCobrador
Dim ColCobrador As New Collection

On Error GoTo Erro_Carrega_Combo_Cobrador

    'Carrega a Coleção de Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then Error 61310

    'Preenche a ComboBox Cobrador com os objetos da coleção de Cobradores
    For Each objCobrador In ColCobrador

        ''If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iCodBanco <> 0 Then
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iInativo = 0 Then
            ComboCobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            ComboCobrador.ItemData(ComboCobrador.NewIndex) = objCobrador.iCodigo
        End If

    Next

    Carrega_Combo_Cobrador = SUCESSO

    Exit Function

Erro_Carrega_Combo_Cobrador:

    Carrega_Combo_Cobrador = Err

    Select Case Err

        Case 61310 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159418)

    End Select

    Exit Function

End Function

Private Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame de Títulos visível
        Frame1(Opcao.SelectedItem.Index).Visible = True

        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False

        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        'Se Frame selecionado foi o de Títulos
        If Opcao.SelectedItem.Index = TAB_SELECAO Then

            Parent.HelpContextID = IDH_EMISSAO_BOLETO_SELECAO
            
            iFramePrincipalAlterado = 0

        'Se Frame selecionado foi o de Parcelas
        ElseIf Opcao.SelectedItem.Index = TAB_EMISSAO Then
            
            Parent.HelpContextID = IDH_EMISSAO_BOLETO_EMISSAO
            
            If iFramePrincipalAlterado <> 0 Then
                
                'Traz as Parcelas para Tela
                lErro = Carrega_Tab_Parcelas()
                If lErro <> SUCESSO Then Error 61311
    
                iFramePrincipalAlterado = 0
            
            End If
            
        End If

    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case Err

        Case 61311 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159419)

    End Select

    Exit Sub

End Sub

Function Carrega_Tab_Parcelas() As Long

Dim lErro As Long
Dim iCobrador As Integer
Dim lBorderoInicial As Long
Dim lBorderoFinal As Long
Dim lClienteInicial As Long
Dim lClienteFinal As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iExibeBoletosImpressos As Integer
Dim objInfoParcRec As ClassInfoParcRec, colInfoParcRec As New Collection
Dim objFilialCliente As New ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim bIgnorar As Boolean, bTemEmail As Boolean

On Error GoTo Erro_Carrega_Tab_Parcelas

    GL_objMDIForm.MousePointer = vbHourglass
    
    If giTipoVersao = VERSAO_FULL Then
        
        If optCobrador.Value = True Then
            'Se o cobrador está em branco --> Erro
            If Len(Trim(ComboCobrador.Text)) = 0 Then gError 61365
            If Len(Trim(BorderoCobrDe.Text)) = 0 And Len(Trim(ClienteInicial.Text)) = 0 And Len(Trim(VencInic.ClipText)) > 0 Then gError 130040
        Else
            '??? ACERTAR CODIGO DE ERRO DEPOIS
            If Len(Trim(ComboBanco.Text)) = 0 Then gError 61366
        End If
        
    End If
    
    'Move os campos da tela para a Memoria
    lErro = Move_Tela_Memoria(iCobrador, lBorderoInicial, lBorderoFinal, lClienteInicial, lClienteFinal, dtDataInicial, dtDataFinal, iExibeBoletosImpressos)
    If lErro <> SUCESSO Then gError 61312
    
    Set gcolInfoParcRec = New Collection
    
    'Limpa o grid
    Call Grid_Limpa(objGridParcelas)
    
    'Limpa os acumuladores
    QtdParcelasSelecionadas.Caption = "0"
    TotalParcelasSelecionadas.Caption = "0,00"
    QtdParcelas.Caption = "0"
    TotalParcelas.Caption = "0,00"
    
    'Preenche a colecao de Tabelas de Acordo com a Selecao passada
    lErro = CF("ParcelasRec_Le_EmissaoBoleta_Sel", iCobrador, lBorderoInicial, lBorderoFinal, lClienteInicial, lClienteFinal, dtDataInicial, dtDataFinal, iExibeBoletosImpressos, colInfoParcRec)
    If lErro <> SUCESSO And lErro <> 61362 Then gError 61313
    
    'Alterado por Wagner 19/02/2009
    'Percorre todas as Parcelas da Coleção e acrescenta a tarifa bancária
    For Each objInfoParcRec In colInfoParcRec
        
        bIgnorar = False
            
        If EmailValido(2).Value = False Then
        
            objFilialCliente.lCodCliente = objInfoParcRec.lCliente
            objFilialCliente.iCodFilial = objInfoParcRec.iFilialCliente
    
            'le o nome reduzido da filial  cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then Error 61315
            
            Set objEndereco = New ClassEndereco
            
            objEndereco.lCodigo = objFilialCliente.lEndereco
        
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then Error 61315

            lErro = CF("Endereco_Trata_Customizacao", objEndereco)
            If lErro <> SUCESSO Then gError 196232
        
            bTemEmail = InStr(1, objEndereco.sEmail, "@") <> 0 And InStr(1, objEndereco.sEmail, ".") <> 0
            
            If EmailValido(0).Value And Not bTemEmail Then bIgnorar = True
            If EmailValido(1).Value And bTemEmail Then bIgnorar = True

        End If

        If Not bIgnorar Then

            gcolInfoParcRec.Add objInfoParcRec
            objInfoParcRec.dValor = objInfoParcRec.dValor + StrParaDbl(ValorTafBanc.Text)
            
        End If
    
    Next
    
    'Não encontrou Parcelas
    If lErro = 61362 Then gError 61363
    
    'Preenche o Grid Com as Parcelas Encontradas
    lErro = Grid_Parcelas_Preenche()
    If lErro <> SUCESSO Then gError 61314
        
    GL_objMDIForm.MousePointer = vbDefault
        
    Carrega_Tab_Parcelas = SUCESSO
        
    Exit Function
    
Erro_Carrega_Tab_Parcelas:

    GL_objMDIForm.MousePointer = vbDefault

    Carrega_Tab_Parcelas = gErr
    
    Select Case gErr
        
        Case 61312, 61313, 61314 'Tratado na rotina chamada
        
        Case 61363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PARCELAS_REC_SEL", gErr)
        
        Case 61365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
        
        Case 130040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDEROCOBR_NAO_INFORMADO", gErr)

        Case 61366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_INFORMADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159420)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(iCobrador As Integer, lBorderoInicial As Long, lBorderoFinal As Long, lClienteInicial As Long, lClienteFinal As Long, dtDataInicial As Date, dtDataFinal As Date, iExibeBoletosImpressos As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    
    If giTipoVersao = VERSAO_FULL Then
        
        'Guarda o cobrador
        iCobrador = Codigo_Extrai(ComboCobrador.Text)

        'Se Preencheu o Bordero Inicial
        If Len(Trim(BorderoCobrDe.Text)) > 0 Then
            lBorderoInicial = CLng(BorderoCobrDe.Text)
        Else
            lBorderoInicial = 0
        End If

        'Se Preencheu o bordero Final
        If Len(Trim(BorderoCobrAte.Text)) > 0 Then
            lBorderoFinal = CLng(BorderoCobrAte.Text)
        Else
            lBorderoFinal = 0
        End If
    
    End If
    
    'Se Preencheu o Cliente Inicial
    If Len(Trim(ClienteInicial.Text)) > 0 Then
        lClienteInicial = LCodigo_Extrai(ClienteInicial.Text)
    Else
        lClienteInicial = 0
    End If

    'Se Preencheu o Cliente Final
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        lClienteFinal = LCodigo_Extrai(ClienteFinal.Text)
    Else
        lClienteFinal = 0
    End If

    'Se o Vencimento Inicial foi preenchido
    If Len(Trim(VencInic.ClipText)) > 0 Then
        dtDataInicial = CDate(VencInic.Text)
    Else
        dtDataInicial = DATA_NULA
    End If

    'Se o Vencimento Final foi preenchido
    If Len(Trim(VencFim.ClipText)) > 0 Then
        dtDataFinal = CDate(VencFim.Text)
    Else
        dtDataFinal = DATA_NULA
    End If

    If ComBoletosImpressos.Value = vbChecked Then
        iExibeBoletosImpressos = vbChecked
    ElseIf ComBoletosImpressos.Value = vbUnchecked Then
        iExibeBoletosImpressos = vbUnchecked
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159421)

    End Select

    Exit Function

End Function

Private Function Grid_Parcelas_Preenche() As Long
'Preenche o Grid Parcelas com os dados de gcolInfoParcRec

Dim iLinha As Integer
Dim objInfoParcRec As ClassInfoParcRec
Dim objFilialEmpresa As New AdmFiliais
Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_Grid_Parcelas_Preenche

    Call Grid_Limpa(objGridParcelas)
    
    'Se o número de parcelas for maior que o número de linhas do Grid
    If gcolInfoParcRec.Count + 1 > GridParcelas.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de parcelas
        GridParcelas.Rows = gcolInfoParcRec.Count + 1
        
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridParcelas)

    End If

    iLinha = 0

    'Percorre todas as Parcelas da Coleção
    For Each objInfoParcRec In gcolInfoParcRec

        iLinha = iLinha + 1
        
        dTotal = dTotal + objInfoParcRec.dValor
        
        'Passa para a tela os dados da Parcela em questão
        If objInfoParcRec.dtVencimento <> DATA_NULA Then GridParcelas.TextMatrix(iLinha, iGrid_DataVencimento_Col) = objInfoParcRec.dtVencimento
        GridParcelas.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcRec.sSiglaDocumento
        GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col) = objInfoParcRec.lNumTitulo
        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcRec.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_ValorCobrado_Col) = Format(objInfoParcRec.dValor, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col) = Format(objInfoParcRec.sNomeRedCliente, "Standard")
        
        If giTipoVersao = VERSAO_FULL Then
            
            If optCobrador.Value = True Then
                GridParcelas.TextMatrix(iLinha, iGrid_Cobrador_Col) = ComboCobrador.Text
                
                'preenche o objFilialEmpresa
                objFilialEmpresa.iCodFilial = objInfoParcRec.iFilialEmpresa
            
                'le o Nome da Filial
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO And lErro <> 27378 Then Error 61315
            
                'Se não encontrou a Filial Empresa
                If lErro = 27378 Then Error 61316
            
                GridParcelas.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
            Else
                GridParcelas.TextMatrix(iLinha, iGrid_Cobrador_Col) = COBRADOR_PROPRIA_EMPRESA & SEPARADOR & "Própria Empresa"
            End If
            
        ElseIf giTipoVersao = VERSAO_LIGHT Then
            GridParcelas.TextMatrix(iLinha, iGrid_Cobrador_Col) = COBRADOR_PROPRIA_EMPRESA & SEPARADOR & "Própria Empresa"
        End If
        
        If objInfoParcRec.dtDataImpressaoBoleto <> DATA_NULA Then GridParcelas.TextMatrix(iLinha, iGrid_ImpressoEm_Col) = objInfoParcRec.dtDataImpressaoBoleto
        
    Next
    
    'Informa os Totais
    TotalParcelas.Caption = Format(dTotal, "Standard")
    QtdParcelas.Caption = CStr(iLinha)
    
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridParcelas.iLinhasExistentes = iLinha

    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    Grid_Parcelas_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grid_Parcelas_Preenche:

    Grid_Parcelas_Preenche = Err
    
    Select Case Err
        
        Case 61315 'Tratado na rotina chamada
        
        Case 61316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objInfoParcRec.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 159422)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcar_Click()
'Marca todas as parcelas no Grid

Dim iLinha As Integer
Dim dTotalParcelasSelecionadas As Double
Dim iNumParcelasSelecionadas As Integer
Dim objInfoParcRec As ClassInfoParcRec

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridParcelas.iLinhasExistentes

        'Marca na tela a parcela em questão
        GridParcelas.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_CHECADO
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcRec = gcolInfoParcRec.Item(iLinha)
        
        'Marca no Obj a parcela em questão
        objInfoParcRec.iMarcada = SELECIONAR_CHECADO
        
        dTotalParcelasSelecionadas = dTotalParcelasSelecionadas + CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorCobrado_Col))
        iNumParcelasSelecionadas = iNumParcelasSelecionadas + 1
    
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    'Atualiza na tela os campos Qtd de Parcelas selecionadas e Valor total das Parcelas selecionados
    QtdParcelasSelecionadas.Caption = CStr(iNumParcelasSelecionadas)
    TotalParcelasSelecionadas.Caption = CStr(Format(dTotalParcelasSelecionadas, "Standard"))
    
End Sub

Private Sub BotaoDesmarcar_Click()
'Desmarca todas as parcelas marcadas no Grid

Dim iLinha As Integer
Dim objInfoParcRec As New ClassInfoParcRec
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridParcelas.iLinhasExistentes

        'Desmarca na tela a parcela em questão
        GridParcelas.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_NAO_CHECADO
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcRec = gcolInfoParcRec.Item(iLinha)
        
        'Desmarca no Obj a parcela em questão
        objInfoParcRec.iMarcada = SELECIONAR_NAO_CHECADO
        
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    'Limpa na tela os campos Qtd de Parcelas selecionadas e Valor total Parcelas selecionadas
    QtdParcelasSelecionadas.Caption = CStr(0)
    TotalParcelasSelecionadas.Caption = CStr(Format(0, "Standard"))
    
End Sub

Private Sub GridParcelas_Click()
    
Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If
    
End Sub

Private Sub GridParcelas_GotFocus()
    
    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_EnterCell()
    
Dim iAlterado As Integer

    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
    
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)
       
End Sub

Private Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)
    
End Sub

Private Sub Selecionar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)
        
End Sub

Private Sub Selecionar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
        
End Sub

Private Sub Selecionar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Selecionar
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Selecionar_Click()

Dim objInfoParcRec As New ClassInfoParcRec
Dim iLinha As Integer

    iLinha = GridParcelas.Row
    
    If GridParcelas.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_CHECADO Then
        
        TotalParcelasSelecionadas.Caption = Format(StrParaDbl(TotalParcelasSelecionadas.Caption) + StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorCobrado_Col)), "Standard")
        QtdParcelasSelecionadas.Caption = StrParaLong(QtdParcelasSelecionadas.Caption) + 1
    
        Set objInfoParcRec = gcolInfoParcRec.Item(iLinha)
    
        objInfoParcRec.iMarcada = SELECIONAR_CHECADO
    
    ElseIf GridParcelas.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_NAO_CHECADO Then
    
        TotalParcelasSelecionadas.Caption = Format(StrParaDbl(TotalParcelasSelecionadas.Caption) - StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorCobrado_Col)), "Standard")
        QtdParcelasSelecionadas.Caption = StrParaLong(QtdParcelasSelecionadas.Caption) - 1
    
        Set objInfoParcRec = gcolInfoParcRec.Item(iLinha)
    
        objInfoParcRec.iMarcada = SELECIONAR_NAO_CHECADO
        
    End If
        
End Sub

Private Sub VencFim_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VencFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VencFim_Validate

    If Len(VencFim.ClipText) = 0 Then Exit Sub

    'verifica se a data final é válida
    lErro = Data_Critica(VencFim.Text)
    If lErro <> SUCESSO Then Error 61357

    'verifica se a data Inicial está Preenchida
    If Len(VencInic.ClipText) > 0 Then
        
        If CDate(VencInic) > CDate(VencFim.Text) Then Error 61356
    
    End If
    
    Exit Sub

Erro_VencFim_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 61356
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
        
        Case 61357

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159423)

    End Select

    Exit Sub

End Sub

Private Sub VencInic_Change()
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VencFim_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VencFim)

End Sub


Private Sub VencInic_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VencInic)

End Sub

Private Sub VencInic_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VencInic_Validate

    If Len(VencInic.ClipText) = 0 Then Exit Sub

    'verifica se a data final é válida
    lErro = Data_Critica(VencInic.Text)
    If lErro <> SUCESSO Then Error 61357

    'verifica se a data Inicial está preenchida
    If Len(VencFim.ClipText) > 0 Then
        If CDate(VencInic.Text) > CDate(VencFim.Text) Then Error 61356
    End If

    Exit Sub

Erro_VencInic_Validate:

    Cancel = True
    
    Select Case Err

        Case 61356
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
        
        Case 61357

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159424)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EMISSAO_BOLETO_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Emissão de Boletos Bancários"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "EmissaoBoletos"
    
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is BorderoCobrDe Then
            Call LabelBorderoCobrDe_Click
        End If
        
    End If
    
End Sub



Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelBorderoCobrDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelBorderoCobrDe, Source, X, Y)
End Sub

Private Sub LabelBorderoCobrDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelBorderoCobrDe, Button, Shift, X, Y)
End Sub

Private Sub LabelBorderoCobrAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelBorderoCobrAte, Source, X, Y)
End Sub

Private Sub LabelBorderoCobrAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelBorderoCobrAte, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub TotalParcelas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalParcelas, Source, X, Y)
End Sub

Private Sub TotalParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalParcelas, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub QtdParcelas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdParcelas, Source, X, Y)
End Sub

Private Sub QtdParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdParcelas, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub TotalParcelasSelecionadas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalParcelasSelecionadas, Source, X, Y)
End Sub

Private Sub TotalParcelasSelecionadas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalParcelasSelecionadas, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub QtdParcelasSelecionadas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdParcelasSelecionadas, Source, X, Y)
End Sub

Private Sub QtdParcelasSelecionadas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdParcelasSelecionadas, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub ValorTafBanc_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTafBanc_Validate

    'Veifica se ValorTafBanc está preenchida
    If Len(Trim(ValorTafBanc.Text)) <> 0 Then

       'Critica a ValorTafBanc
       lErro = Valor_Positivo_Critica(ValorTafBanc.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If

    Exit Sub

Erro_ValorTafBanc_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Private Sub ValorTafBanc_Change()
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IncluirCobrBanc_Click()
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
    If IncluirCobrBanc.Value = vbChecked Then
        ValorTafBanc.Enabled = True
    Else
        ValorTafBanc.Text = ""
        ValorTafBanc.Enabled = False
    End If
End Sub

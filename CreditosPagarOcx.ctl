VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl CreditosPagarOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4995
      Index           =   1
      Left            =   150
      TabIndex        =   0
      Top             =   750
      Width           =   9195
      Begin VB.Frame Frame6 
         Caption         =   "Situação Atual"
         Height          =   765
         Left            =   195
         TabIndex        =   94
         Top             =   4215
         Width           =   8490
         Begin VB.CommandButton BotaoBaixas 
            Caption         =   "Baixas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3330
            TabIndex        =   95
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
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
            Left            =   1020
            TabIndex        =   97
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Saldo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   285
            Left            =   1725
            TabIndex        =   96
            Top             =   308
            Width           =   1530
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   "Dados Principais"
         Height          =   1515
         Left            =   195
         TabIndex        =   42
         Top             =   105
         Width           =   8475
         Begin VB.ComboBox Tipo 
            Height          =   315
            Left            =   1515
            Sorted          =   -1  'True
            TabIndex        =   3
            Text            =   " "
            Top             =   690
            Width           =   2775
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5820
            TabIndex        =   2
            Top             =   285
            Width           =   1815
         End
         Begin MSMask.MaskEdBox NumTitulo 
            Height          =   300
            Left            =   5835
            TabIndex        =   4
            Top             =   690
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1515
            TabIndex        =   1
            Top             =   285
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   2565
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1095
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1530
            TabIndex        =   5
            Top             =   1095
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotal 
            Height          =   300
            Left            =   5820
            TabIndex        =   6
            Top             =   1095
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label LabelTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
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
            Left            =   990
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   750
            Width           =   450
         End
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   5175
            TabIndex        =   51
            Top             =   1155
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
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
            Left            =   675
            TabIndex        =   52
            Top             =   1155
            Width           =   765
         End
         Begin VB.Label NumeroLabel 
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
            Left            =   4995
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   53
            Top             =   750
            Width           =   720
         End
         Begin VB.Label FornecedorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
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
            Left            =   405
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   54
            Top             =   345
            Width           =   1035
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   " Filial:"
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
            Left            =   5160
            TabIndex        =   55
            Top             =   345
            Width           =   525
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Valores"
         Height          =   2520
         Left            =   195
         TabIndex        =   44
         Top             =   1665
         Width           =   8475
         Begin VB.Frame Frame2 
            Caption         =   "Retenções"
            Height          =   1035
            Left            =   3420
            TabIndex        =   84
            Top             =   1395
            Width           =   4920
            Begin MSMask.MaskEdBox ValorIRRF 
               Height          =   300
               Left            =   1110
               TabIndex        =   85
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox PISRetido 
               Height          =   300
               Left            =   3105
               TabIndex        =   87
               Top             =   270
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox COFINSRetido 
               Height          =   300
               Left            =   1110
               TabIndex        =   89
               Top             =   615
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox CSLLRetido 
               Height          =   300
               Left            =   3105
               TabIndex        =   91
               Top             =   645
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Caption         =   "CSLL:"
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
               Left            =   2535
               TabIndex        =   92
               Top             =   690
               Width           =   525
            End
            Begin VB.Label Label2 
               Caption         =   "COFINS:"
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
               Left            =   270
               TabIndex        =   90
               Top             =   660
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "PIS:"
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
               Left            =   2685
               TabIndex        =   88
               Top             =   315
               Width           =   375
            End
            Begin VB.Label Label16 
               Caption         =   "IR:"
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
               Left            =   765
               TabIndex        =   86
               Top             =   285
               Width           =   300
            End
         End
         Begin VB.Frame Frame3 
            Height          =   630
            Left            =   165
            TabIndex        =   45
            Top             =   1455
            Width           =   3060
            Begin VB.CheckBox DebitoIPI 
               Caption         =   "Débito"
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
               Left            =   2040
               TabIndex        =   15
               Top             =   285
               Width           =   885
            End
            Begin MSMask.MaskEdBox ValorIPI 
               Height          =   300
               Left            =   690
               TabIndex        =   14
               Top             =   225
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label7 
               Caption         =   "IPI:"
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
               Left            =   285
               TabIndex        =   56
               Top             =   285
               Width           =   315
            End
         End
         Begin VB.Frame Frame4 
            Height          =   645
            Left            =   180
            TabIndex        =   46
            Top             =   210
            Width           =   5895
            Begin VB.CheckBox DebitoICMS 
               Caption         =   "Débito"
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
               Left            =   4890
               TabIndex        =   9
               Top             =   270
               Width           =   900
            End
            Begin MSMask.MaskEdBox ValorICMS 
               Height          =   300
               Left            =   675
               TabIndex        =   7
               Top             =   210
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox ValorICMSSubst 
               Height          =   300
               Left            =   3480
               TabIndex        =   8
               Top             =   210
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "ICMS Subst:"
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
               Left            =   2310
               TabIndex        =   57
               Top             =   270
               Width           =   1065
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "ICMS:"
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
               Left            =   90
               TabIndex        =   58
               Top             =   270
               Width           =   525
            End
         End
         Begin MSMask.MaskEdBox ValorProdutos 
            Height          =   300
            Left            =   7095
            TabIndex        =   10
            Top             =   405
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   300
            Left            =   855
            TabIndex        =   11
            Top             =   1057
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox OutrasDespesas 
            Height          =   300
            Left            =   7095
            TabIndex        =   13
            Top             =   1057
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   300
            Left            =   3660
            TabIndex        =   12
            Top             =   1057
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label19 
            Caption         =   "Seguro:"
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
            Left            =   2865
            TabIndex        =   59
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label Label18 
            Caption         =   "Outras Despesas:"
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
            Left            =   5415
            TabIndex        =   60
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Frete:"
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
            Left            =   270
            TabIndex        =   61
            Top             =   1080
            Width           =   525
         End
         Begin VB.Label Label20 
            Caption         =   "Produtos:"
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
            Height          =   255
            Left            =   6180
            TabIndex        =   62
            Top             =   428
            Width           =   840
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5070
      Index           =   2
      Left            =   270
      TabIndex        =   16
      Top             =   825
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4440
         TabIndex        =   93
         Tag             =   "1"
         Top             =   1560
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7725
         TabIndex        =   21
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   900
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6300
         TabIndex        =   20
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6270
         TabIndex        =   22
         Top             =   375
         Width           =   2700
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   30
         Top             =   1920
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   32
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   31
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   34
         Top             =   1500
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   47
         Top             =   3450
         Width           =   5895
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            Left            =   240
            TabIndex        =   63
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label CTBLabel7 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
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
            Left            =   1125
            TabIndex        =   64
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   65
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   66
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
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
         Left            =   3450
         TabIndex        =   25
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   26
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3450
         TabIndex        =   29
         Top             =   1890
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   28
         Top             =   1830
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   27
         Top             =   1875
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
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
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   19
         Top             =   525
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5580
         TabIndex        =   18
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3810
         TabIndex        =   17
         Top             =   120
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   0
         TabIndex        =   33
         Top             =   1185
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2985
         Left            =   6330
         TabIndex        =   35
         Top             =   1485
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6330
         TabIndex        =   36
         Top             =   1485
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Height          =   195
         Left            =   6330
         TabIndex        =   23
         Top             =   690
         Width           =   690
      End
      Begin VB.Label CTBLabel21 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   45
         TabIndex        =   67
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   68
         Top             =   120
         Width           =   1530
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   4230
         TabIndex        =   69
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   70
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   71
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
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
         Left            =   1995
         TabIndex        =   72
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         Left            =   45
         TabIndex        =   73
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   74
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabelContas 
         Caption         =   "Plano de Contas"
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
         Left            =   6345
         TabIndex        =   75
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6345
         TabIndex        =   76
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabelTotais 
         Caption         =   "Totais:"
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
         Height          =   225
         Left            =   1800
         TabIndex        =   77
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   78
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   79
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBLabel8 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   45
         TabIndex        =   80
         Top             =   555
         Width           =   480
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   81
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   82
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6120
      ScaleHeight     =   495
      ScaleWidth      =   3165
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   90
      Width           =   3225
      Begin VB.CommandButton BotaoDocOriginal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   60
         Picture         =   "CreditosPagarOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Consulta de nota fiscal"
         Top             =   60
         Width           =   1005
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   390
         Left            =   1147
         Picture         =   "CreditosPagarOcx.ctx":0F0A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   390
         Left            =   1649
         Picture         =   "CreditosPagarOcx.ctx":1064
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   390
         Left            =   2151
         Picture         =   "CreditosPagarOcx.ctx":11EE
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   390
         Left            =   2655
         Picture         =   "CreditosPagarOcx.ctx":1720
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5505
      Left            =   120
      TabIndex        =   43
      Top             =   420
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9710
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
Attribute VB_Name = "CreditosPagarOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'Mnemônicos
Private Const FORNECEDOR_COD As String = "Fornecedor_Codigo"
Private Const FORNECEDOR_NOME As String = "Fornecedor_Nome"
Private Const FILIAL_COD As String = "FilialForn_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const FILIAL_CONTA As String = "FilialForn_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialForn_CGC_CPF"
Private Const NUMERO1 As String = "Numero_Nota_Fiscal"
Private Const EMISSAO1 As String = "Data_Emissao"
Private Const VALORTOTAL1 As String = "Valor_Total"
Private Const VALOR_ICMS As String = "Valor_ICMS"
Private Const VALOR_ICMS_SUBST As String = "Valor_ICMS_Subst"
Private Const DEBITA_ICMS As String = "Debita_ICMS"
Private Const VALOR_IR As String = "Valor_IRRF"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_IPI As String = "Valor_IPI"
Private Const DEBITA_IPI As String = "Debita_IPI"
Private Const OUTRAS_DESPESAS As String = "Valor_OutrasDesp"
Private Const TIPO1 As String = "Tipo"
Private Const CONTA_DESP_ESTOQUE As String = "Conta_Desp_Estoque"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"

'Constantes globais da Tela
Dim iFrameAtual As Integer
Public iAlterado As Integer
Dim iFornecedorAlterado As Integer
Private glNumIntDoc As Long

'Eventos de Browse
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Contabilizacao = 2

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCreditoPagar As New ClassCreditoPagar

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CreditosPagForn"

    'Lê os dados da Tela CreditoPagar
    lErro = Move_Tela_Memoria(objCreditoPagar)
    If lErro <> SUCESSO Then Error 17372

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Fornecedor", objCreditoPagar.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objCreditoPagar.iFilial, 0, "Filial"
    colCampoValor.Add "SiglaDocumento", objCreditoPagar.sSiglaDocumento, STRING_CRED_PAG_SIGLA, "SiglaDocumento"
    colCampoValor.Add "NumTitulo", objCreditoPagar.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objCreditoPagar.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorTotal", objCreditoPagar.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorICMS", objCreditoPagar.dValorICMS, 0, "ValorICMS"
    colCampoValor.Add "ValorICMSSubst", objCreditoPagar.dValorICMS, 0, "ValorICMSSubst"
    colCampoValor.Add "DebitoICMS", objCreditoPagar.iDebitoICMS, 0, "DebitoICMS"
    colCampoValor.Add "ValorProdutos", objCreditoPagar.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorIRRF", objCreditoPagar.dValorIRRF, 0, "ValorIRRF"
    colCampoValor.Add "ValorFrete", objCreditoPagar.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorSeguro", objCreditoPagar.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "OutrasDespesas", objCreditoPagar.dOutrasDespesas, 0, "OutrasDespesas"
    colCampoValor.Add "ValorIPI", objCreditoPagar.dValorIPI, 0, "ValorIPI"
    colCampoValor.Add "DebitoIPI", objCreditoPagar.iDebitoIPI, 0, "DebitoIPI"
    colCampoValor.Add "NumIntDoc", objCreditoPagar.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "PISRetido", objCreditoPagar.dPISRetido, 0, "PISRetido"
    colCampoValor.Add "COFINSRetido", objCreditoPagar.dCOFINSRetido, 0, "COFINSRetido"
    colCampoValor.Add "CSLLRetido", objCreditoPagar.dCSLLRetido, 0, "CSLLRetido"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, objCreditoPagar.iFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 17372 'Tratada na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155249)

    End Select

    Exit Sub

End Sub

Public Function Move_Tela_Memoria(objCreditoPagar As ClassCreditoPagar) As Long
'Lê dados que estão na tela e coloca em objCreditoPagar

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    objCreditoPagar.iFilialEmpresa = giFilialEmpresa

    'Verifica se campos da tela estão preenchidos
    If Len(Trim(Fornecedor.Text)) > 0 Then
        
        'Lê Fornecedor
        objFornecedor.sNomeReduzido = Fornecedor.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 17370
        
        'Não encontrou o nome reduzido do fornecedor --> erro
        If lErro = 6681 Then Error 17371

        objCreditoPagar.lFornecedor = objFornecedor.lCodigo

    End If

    If Len(Trim(Filial.Text)) > 0 Then
        objCreditoPagar.iFilial = Codigo_Extrai(Filial.Text)
    End If

    objCreditoPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    
    If Len(Trim(NumTitulo.Text)) > 0 Then
       objCreditoPagar.lNumTitulo = CLng(NumTitulo.Text)
    Else
        objCreditoPagar.lNumTitulo = 0
    End If
    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objCreditoPagar.dtDataEmissao = DATA_NULA
    Else
        objCreditoPagar.dtDataEmissao = CDate(DataEmissao.Text)
    End If
    
    If Len(Trim(ValorTotal.Text)) > 0 Then objCreditoPagar.dValorTotal = CDbl(ValorTotal.Text)
    If Len(Trim(ValorSeguro.Text)) > 0 Then objCreditoPagar.dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(ValorFrete.Text)) > 0 Then objCreditoPagar.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(OutrasDespesas.Text)) > 0 Then objCreditoPagar.dOutrasDespesas = CDbl(OutrasDespesas.Text)
    If Len(Trim(ValorProdutos.Text)) > 0 Then objCreditoPagar.dValorProdutos = CDbl(ValorProdutos.Text)
    If Len(Trim(ValorICMS.Text)) > 0 Then objCreditoPagar.dValorICMS = CDbl(ValorICMS.Text)
    If Len(Trim(ValorICMSSubst.Text)) > 0 Then objCreditoPagar.dValorICMSSubst = CDbl(ValorICMSSubst.Text)
    If Len(Trim(ValorIPI.Text)) > 0 Then objCreditoPagar.dValorIPI = CDbl(ValorIPI.Text)
    If Len(Trim(ValorIRRF.Text)) > 0 Then objCreditoPagar.dValorIRRF = CDbl(ValorIRRF.Text)
    If Len(Trim(PISRetido.Text)) > 0 Then objCreditoPagar.dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(COFINSRetido.Text)) > 0 Then objCreditoPagar.dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.Text)) > 0 Then objCreditoPagar.dCSLLRetido = CDbl(CSLLRetido.Text)
    
    objCreditoPagar.iDebitoICMS = DebitoICMS.Value
    objCreditoPagar.iDebitoIPI = DebitoIPI.Value
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 17370 'Tratado na Rotina chamada
        
        Case 17371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155250)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCreditoPagar As New ClassCreditoPagar

On Error GoTo Erro_Tela_Preenche

    objCreditoPagar.sSiglaDocumento = colCampoValor.Item("SiglaDocumento").vValor

    If objCreditoPagar.sSiglaDocumento <> "" Then

        'Carrega objCreditoPagar com os dados passados em colCampoValor
        objCreditoPagar.lFornecedor = colCampoValor.Item("Fornecedor").vValor
        objCreditoPagar.iFilial = colCampoValor.Item("Filial").vValor
        objCreditoPagar.lNumTitulo = colCampoValor.Item("NumTitulo").vValor
        objCreditoPagar.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objCreditoPagar.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objCreditoPagar.dValorICMS = colCampoValor.Item("ValorICMS").vValor
        objCreditoPagar.dValorICMSSubst = colCampoValor.Item("ValorICMSSubst").vValor
        objCreditoPagar.iDebitoICMS = colCampoValor.Item("DebitoICMS").vValor
        objCreditoPagar.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
        objCreditoPagar.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
        objCreditoPagar.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        objCreditoPagar.dValorIPI = colCampoValor.Item("ValorIPI").vValor
        objCreditoPagar.iDebitoIPI = colCampoValor.Item("DebitoIPI").vValor
        objCreditoPagar.dValorFrete = colCampoValor.Item("ValorFrete").vValor
        objCreditoPagar.dOutrasDespesas = colCampoValor.Item("OutrasDespesas").vValor
        objCreditoPagar.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
        objCreditoPagar.dPISRetido = colCampoValor.Item("PISRetido").vValor
        objCreditoPagar.dCOFINSRetido = colCampoValor.Item("COFINSRetido").vValor
        objCreditoPagar.dCSLLRetido = colCampoValor.Item("CSLLRetido").vValor
        
        'Verifica se o credito a pagar existe
        lErro = CF("CreditoPagar_Le", objCreditoPagar)
        If lErro <> SUCESSO And lErro <> 17071 Then Error 17073
        
        lErro = Exibe_Dados_CreditoPagar(objCreditoPagar)
        If lErro <> SUCESSO Then Error 17373

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 17373 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155251)

    End Select

    Exit Sub

End Sub

Function Exibe_Dados_CreditoPagar(objCreditoPagar As ClassCreditoPagar) As Long
'Exibe os dados de Crédito Pagar na tela

Dim lErro As Long, bCancel As Boolean
Dim objCodigoNome As New AdmlCodigoNome

On Error GoTo Erro_Exibe_Dados_CreditoPagar

    objCodigoNome.lCodigo = objCreditoPagar.lFornecedor

    'Lê o NomeReduzido de Fornecedor
    lErro = CF("Fornecedor_Le_NomeRed", objCodigoNome)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 17374
    
    'Fornecedor não existe
    If lErro = 6681 Then Error 33994
        
    'Preenche campos da tela
    Fornecedor.Text = objCodigoNome.sNome
    Call Fornecedor_Validate(bCancel)

    Filial.Text = CStr(objCreditoPagar.iFilial)
    Call Filial_Validate(bCancel)

    Tipo.Text = objCreditoPagar.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)
    
    If objCreditoPagar.lNumTitulo = 0 Then
        NumTitulo.PromptInclude = False
        NumTitulo.Text = ""
        NumTitulo.PromptInclude = True
    Else
        NumTitulo.PromptInclude = False
        NumTitulo.Text = CStr(objCreditoPagar.lNumTitulo)
        NumTitulo.PromptInclude = True
    End If

    Call DateParaMasked(DataEmissao, objCreditoPagar.dtDataEmissao)

    If objCreditoPagar.dValorTotal <> 0 Then
        ValorTotal.Text = Format(objCreditoPagar.dValorTotal, "Fixed")
    Else
        ValorTotal.Text = ""
    End If
    
    If objCreditoPagar.dValorICMS <> 0 Then
        ValorICMS.Text = Format(objCreditoPagar.dValorICMS, "Fixed")
    Else
        ValorICMS.Text = ""
    End If
    
    If objCreditoPagar.dValorICMSSubst <> 0 Then
        ValorICMSSubst.Text = Format(objCreditoPagar.dValorICMSSubst, "Fixed")
    Else
        ValorICMSSubst.Text = ""
    End If
        
    If objCreditoPagar.dValorProdutos <> 0 Then
        ValorProdutos.Text = Format(objCreditoPagar.dValorProdutos, "Fixed")
    Else
        ValorProdutos.Text = ""
    End If

    If objCreditoPagar.dValorIRRF <> 0 Then
        ValorIRRF.Text = Format(objCreditoPagar.dValorIRRF, "Fixed")
    Else
        ValorIRRF.Text = ""
    End If
    
    If objCreditoPagar.dPISRetido <> 0 Then
        PISRetido.Text = Format(objCreditoPagar.dPISRetido, "Fixed")
    Else
        PISRetido.Text = ""
    End If
    
    If objCreditoPagar.dCOFINSRetido <> 0 Then
        COFINSRetido.Text = Format(objCreditoPagar.dCOFINSRetido, "Fixed")
    Else
        COFINSRetido.Text = ""
    End If
    
    If objCreditoPagar.dCSLLRetido <> 0 Then
        CSLLRetido.Text = Format(objCreditoPagar.dCSLLRetido, "Fixed")
    Else
        CSLLRetido.Text = ""
    End If
    
    If objCreditoPagar.dValorFrete <> 0 Then
        ValorFrete.Text = Format(objCreditoPagar.dValorFrete, "Fixed")
    Else
        ValorFrete.Text = ""
    End If
        
    If objCreditoPagar.dValorSeguro <> 0 Then
        ValorSeguro.Text = Format(objCreditoPagar.dValorSeguro, "Fixed")
    Else
        ValorSeguro.Text = ""
    End If
        
    If objCreditoPagar.dOutrasDespesas <> 0 Then
        OutrasDespesas.Text = Format(objCreditoPagar.dOutrasDespesas, "Fixed")
    Else
        OutrasDespesas.Text = ""
    End If
        
    If objCreditoPagar.dValorIPI <> 0 Then
        ValorIPI.Text = Format(objCreditoPagar.dValorIPI, "Fixed")
    Else
        ValorIPI.Text = ""
    End If
        
    DebitoICMS.Value = objCreditoPagar.iDebitoICMS
    DebitoIPI.Value = objCreditoPagar.iDebitoIPI
    
    Saldo.Caption = Format(objCreditoPagar.dSaldo, "STANDARD")
    glNumIntDoc = objCreditoPagar.lNumIntDoc
    
    'Exibe os dados contábeis na tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objCreditoPagar.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then Error 39600

    iAlterado = 0

    Exibe_Dados_CreditoPagar = SUCESSO

Exit Function

Erro_Exibe_Dados_CreditoPagar:

    Exibe_Dados_CreditoPagar = Err

    Select Case Err

        Case 17374, 17375, 39600 'Tratado na Rotina chamada
        
        Case 33994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objCodigoNome.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155252)

    End Select

    Exit Function

End Function

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim lNumIntNF As Long
Dim objFornecedor As New ClassFornecedor
Dim objCreditoPagar As New ClassCreditoPagar
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoDocOriginal_Click
    
    'Critica se os Campos estão Preenchidos(Fornecedor, Filial, Tipo, Número, DataEmissao)
    lErro = DocOriginal_Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 79635

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79636

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 79637
    
    'Preenche objCreditoPagar
    objCreditoPagar.lFornecedor = objFornecedor.lCodigo
    objCreditoPagar.iFilial = Codigo_Extrai(Filial.Text)
    objCreditoPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objCreditoPagar.lNumTitulo = CLng(NumTitulo.Text)
    objCreditoPagar.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objCreditoPagar.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objCreditoPagar.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Crédito (Baixados ou não)
    lErro = CF("CreditoPagar_Le_Numero", objCreditoPagar)
    If lErro <> SUCESSO And lErro <> 17172 Then gError 79638
    
    'Se não encontrou o título => erro
    If lErro = 17172 Then gError 79640

    'Se o documento é do tipo NCP => erro (não possui doc original)
    If objCreditoPagar.sSiglaDocumento = SIGLA_NOTA_CREDITO_PAGAR Then gError 79639
    
    'Verifica se o Crédito foi gerado através do CP ou através de uma NFiscal do FAT
    lErro = CF("DocumentoCPR_OrigemNFiscal", objCreditoPagar.sSiglaDocumento, objCreditoPagar.lNumIntDoc, lNumIntNF)
    If lErro <> SUCESSO And lErro <> 41542 Then gError 79642
    
    'Se não encontrou NFiscal de origem
    If lErro = 41542 Then gError 79643
    
    objNFiscal.lNumIntDoc = lNumIntNF
        
    'Se a versão utilizada for a Full
    If giTipoVersao = VERSAO_FULL Then
        'Chama a Tela 'Entrada - Nota Fiscal de Devolução'
        Call Chama_Tela("NFiscalDev", objNFiscal)
    
    'Se a versão utilizada for a Light
    ElseIf giTipoVersao = VERSAO_LIGHT Then
        
        'Chama a Tela 'Notas Fiscais de Entrada'
        Call Chama_Tela("NFiscalFatura", objNFiscal)
    End If
    
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 79635, 79636, 79638, 79642
        
        Case 79637
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 79639, 79643
           Call Rotina_Erro(vbOKOnly, "ERRO_CREDITO_PAGAR_SEM_DOC_ORIGINAL", gErr, objCreditoPagar.lNumTitulo)
        
        Case 79640
            'Se a data de emissão não está preenchida
            If objCreditoPagar.dtDataEmissao = DATA_NULA Then
                'Não passa a data como parâmetro
                Call Rotina_Erro(vbOKOnly, "ERRO_CREDITOPAGAR_NAO_CADASTRADO2", gErr, objCreditoPagar.lNumTitulo, objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.sSiglaDocumento, "")
            'Senão
            Else
                'A data de emissão é passada como parâmetro
                Call Rotina_Erro(vbOKOnly, "ERRO_CREDITOPAGAR_NAO_CADASTRADO2", gErr, objCreditoPagar.lNumTitulo, objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.sSiglaDocumento, objCreditoPagar.dtDataEmissao)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155253)

    End Select

    Exit Sub

End Sub

Private Function DocOriginal_Critica_CamposPreenchidos() As Long

On Error GoTo Erro_DocOriginal_Critica_CamposPreenchidos

    'Se o Fornecedor não foi preenchido => erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 79631
    
    'Se a Filial não foi selecionada => erro
    If Len(Trim(Filial.Text)) = 0 Then gError 79632
    
    'Se o tipo do Documento não foi selecionado => erro
    If Len(Trim(Tipo.Text)) = 0 Then gError 79633
    
    'Se o número do Documento não foi preenchido => erro
    If Len(Trim(NumTitulo.Text)) = 0 Then gError 79634
    
'    'Se a data de emissão não foi preenchida => erro
'    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79641
    
    DocOriginal_Critica_CamposPreenchidos = SUCESSO
    
    Exit Function
    
Erro_DocOriginal_Critica_CamposPreenchidos:

    DocOriginal_Critica_CamposPreenchidos = gErr
    
    Select Case gErr
    
        Case 79631
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 79632
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79633
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
        Case 79634
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
'        Case 79641
'            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_OBRIGATORIA_DOC_ORIGINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155254)
        
    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCreditoPagar As New ClassCreditoPagar
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colCodigoNome As New AdmColCodigoNome
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 17162

    'Verifica preenchimento de Filial
    If Len(Trim(Filial.Text)) = 0 Then Error 17163

    'Verifica preenchimento de NumTítulo
    If Len(Trim(NumTitulo.Text)) = 0 Then Error 17165

    'Verifica preenchimento do Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Error 17166
    
    'Lê o Fornecedor
    objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 17352
    
    'Não encontrou o Fornecedor
    If lErro = 6681 Then Error 17353
    
   'Preenche objCreditoPagar
    objCreditoPagar.iFilialEmpresa = giFilialEmpresa
    objCreditoPagar.lFornecedor = objFornecedor.lCodigo
    objCreditoPagar.iFilial = Codigo_Extrai(Filial.Text)
    objCreditoPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objCreditoPagar.lNumTitulo = CLng(NumTitulo.Text)

   'Preenche Data Emissão
    If Len(Trim(DataEmissao.ClipText)) = 0 Then
       objCreditoPagar.dtDataEmissao = DATA_NULA
    Else
      objCreditoPagar.dtDataEmissao = CDate(DataEmissao.Text)
    End If

    'Lê Credito Pagar
    lErro = CF("CreditoPagar_Le_Numero", objCreditoPagar)
    If lErro <> SUCESSO And lErro <> 17172 And lErro <> 17379 Then Error 17211
    
    'Não encontrou o Crédito a Pagar ==> Erro
    If lErro = 17172 Then Error 28359
    
    'O Crédito a Pagar está baixado ==> Erro
    If objCreditoPagar.iStatus = STATUS_BAIXADO Then Error 28360
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CREDITOPAGAR", objCreditoPagar.sSiglaDocumento, objCreditoPagar.lNumTitulo)

    If vbMsgRes = vbYes Then

        'Chama CreditoPagar_Exclui(inclusive dados contábeis)
        lErro = CF("CreditoPagar_Exclui", objCreditoPagar, objContabil)
        If lErro <> SUCESSO Then Error 17167

        'Limpa a tela
        Call Limpa_Tela_CreditoPagar

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17167, 17211, 17352 'Tratado na rotina chamada

        Case 17162
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 17163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 17165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case 17166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", Err)

        Case 17353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case 28359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CREDITOPAGAR_NAO_CADASTRADO", Err, objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.sSiglaDocumento, objCreditoPagar.lNumTitulo, objCreditoPagar.dtDataEmissao)
        
        Case 28360
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CREDPAG_BAIXADO", Err, objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.sSiglaDocumento, objCreditoPagar.lNumTitulo, objCreditoPagar.dtDataEmissao)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155255)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 17087

    'Limpa a tela CreditosPagar
    Call Limpa_Tela_CreditoPagar

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 17087
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155256)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se algum campo foi alterado
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17086

    'Limpa a tela CreditosPagar
    Call Limpa_Tela_CreditoPagar

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 17086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155257)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a Data de Emissão está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Verifica se a data é válida
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 17082

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 17082

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155258)

    End Select

    Exit Sub

End Sub

Private Sub DebitoICMS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DebitoIPI_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

   iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoTipo = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
 
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim sFornecedor As String
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub
    
    'Verifica se é uma filial selecionada
    If Filial.ListIndex >= 0 Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 33995
    
    'Se não encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 33996

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo
        
        'Pesquisa se existe filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 33997

        If lErro = 18272 Then
        
            objFornecedor.sNomeReduzido = sFornecedor
            
            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then Error 58662
            
            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            
            'Sugere cadastrar nova Filial
            Error 33998
        
        End If
        
        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome
        
    End If
    
    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 33999
    
    Exit Sub
    
Erro_Filial_Validate:
    
    Cancel = True
        
    Select Case Err
    
        Case 33995, 33997, 58662 'Tratado na rotina chamada
        
        Case 33996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
        Case 33998
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If
        
        Case 33999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155259)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    iFornecedorAlterado = 1
    iAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Lê o Fornecedor
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 17297

            'Lê filiais de Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then Error 17298

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            If iCodFilial = FILIAL_MATRIZ Then
                Filial.ListIndex = 0
            Else
                Call CF("Filial_Seleciona", Filial, iCodFilial)
            End If

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            'Limpa ComboBox Filiais
            Filial.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case Err

        Case 17297, 17298 'TRatado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155260)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_FornecedorLabel_Click

    'Preenche Crédito a Pagar com nome reduzido de Fornecedor
    objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

Erro_FornecedorLabel_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155261)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New colTipoDocumento
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Visible = False
    
    End If
    
    iFrameAtual = 1

    Set objGrid1 = New AdmGrid

    'Tela em questão
    Set objGrid1.objForm = Me

    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoTipo = New AdmEvento

    iAlterado = 0

    'Preenche a ComboBox com  os Tipos de Documentos existentes no BD
    lErro = CF("TiposDocumento_Le_CredPagar", colTipoDocumento)
    If lErro <> SUCESSO Then Error 17063

    For Each objTipoDocumento In colTipoDocumento

        'Preenche a ComboBox Tipo com os objetos da colecao colTipoDocumento
        Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida

    Next

    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then Error 39599
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 17063, 39599 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155262)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39601

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39602

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
        
        Case 39601 'Tratado na rotina chamada

        Case 39602
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub NumeroLabel_Click()

Dim objCreditoPagar As New ClassCreditoPagar
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click

    'Verifica se campos Fornecedor, Filial, Tipo estão preenchidos
    If Len(Trim(Fornecedor.Text)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(Tipo.Text)) = 0 Then Error 17199

    'Move os dados da Tela para objCreditoPagar
    lErro = Move_Tela_Memoria(objCreditoPagar)
    If lErro <> SUCESSO Then Error 17203

    'Armazena fitro Fornecedor, Filial, Tipo para Browse
    colSelecao.Add objCreditoPagar.lFornecedor
    colSelecao.Add objCreditoPagar.iFilial
    colSelecao.Add objCreditoPagar.sSiglaDocumento

    Call Chama_Tela("CreditosPagarLista", colSelecao, objCreditoPagar, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 17199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_CREDITO_PAGAR_NAO_PREENCHIDOS", Err)

        Case 17203 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155263)

    End Select

    Exit Sub

End Sub

Private Sub NumTitulo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumTitulo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCreditoPagar As ClassCreditoPagar

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objCreditoPagar = obj1

    'Exibe os dados na Tela
    lErro = Exibe_Dados_CreditoPagar(objCreditoPagar)
    If lErro <> SUCESSO Then Error 17204

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 17204 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155264)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedor As ClassFornecedor, Cancel As Boolean

On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1

    Fornecedor.Text = objFornecedor.sNomeReduzido
    
    Call Fornecedor_Validate(Cancel)
    
    Me.Show

    Exit Sub

Erro_objEventoFornecedor_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155265)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_DEVOL_CREDITOS_ID
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_DEVOL_CREDITO_CONTABILIZACAO
                        
        End Select

    End If

End Sub

Function Trata_Parametros(Optional objCreditoPagar As ClassCreditoPagar) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um tipo de documento selecionado, exibir seus dados
    If Not (objCreditoPagar Is Nothing) Then
        
        'Verifica se o credito a pagar existe
        lErro = CF("CreditoPagar_Le", objCreditoPagar)
        If lErro <> SUCESSO And lErro <> 17071 Then Error 17073
        
        'Crédito à Pagar não existe
        If lErro = 17071 Then Error 17150

        'Exibe os dados na Tela
        lErro = Exibe_Dados_CreditoPagar(objCreditoPagar)
        If lErro <> SUCESSO Then Error 17208

    End If

    Trata_Parametros = SUCESSO

    iAlterado = 0

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 17073, 17208 'Tratado na rotina chamada
        
        Case 17150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CREDITOPAGAR_NAO_CADASTRADO1", Err, objCreditoPagar.lNumIntDoc)
            Call Limpa_Tela_CreditoPagar

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155266)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub OutrasDespesas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OutrasDespesas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OutrasDespesas_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(OutrasDespesas.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(OutrasDespesas.Text)
    If lErro <> SUCESSO Then Error 17768

    'Põe o valor formatado na tela
    OutrasDespesas.Text = Format(OutrasDespesas.Text, "Fixed")

    Exit Sub

Erro_OutrasDespesas_Validate:

    Cancel = True


    Select Case Err

        Case 17768

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155267)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Click()

   iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Error_Tipo_Validate

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Tipo
    If Tipo.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = CF("SCombo_Seleciona", Tipo)
    If lErro <> SUCESSO And lErro <> 60483 Then Error 17129

    'Se nao encontrar -> Erro
    If lErro = 60483 Then Error 17330

    Exit Sub

Error_Tipo_Validate:

    Cancel = True


    Select Case Err

        Case 17129

        Case 17330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", Err, Tipo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155268)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick
    
    'Diminui data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 17083

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 17083 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155269)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick
    
    'Aumenta data
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 17084

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 17084 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155270)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_CreditoPagar() As Long
'Limpa todos os campos de input da tela CreditosPagar

Dim iIndice As Integer
Dim lErro As Long

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função generica que limpa campos da tela
    Call Limpa_Tela(Me)

    'Limpa os campos que não são limpos pela função acima
    Filial.Clear
    Tipo.Text = ""
    DebitoICMS.Value = vbUnchecked
    DebitoIPI.Value = vbUnchecked
    
    Saldo.Caption = ""
    glNumIntDoc = 0
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCreditoPagar As New ClassCreditoPagar
Dim dSoma As Double
Dim dValorProdutos As Double
Dim dValorSeguro  As Double
Dim dValorICMSSubst  As Double
Dim dValorFrete As Double
Dim dOutrasDespesas As Double
Dim dValorIPI As Double

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 17102

    'Verifica preenchimento de Filial
    If Len(Trim(Filial.Text)) = 0 Then Error 17103

    'Verifica preenchimento do Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Error 17108

    'Verifica preenchimento de NumTítulo
    If Len(Trim(NumTitulo.Text)) = 0 Then Error 17105

    'Verifica preenchimento de Valor
    If Len(Trim(ValorTotal.Text)) = 0 Then Error 17106

    'Verifica preenchimento de Valor Produtos
    If Len(Trim(ValorProdutos.ClipText)) = 0 Then Error 17771

    'Verifica se a soma dos Valores é igual ao Valor Total
    If Len(Trim(ValorProdutos.Text)) > 0 Then dValorProdutos = CDbl(ValorProdutos.Text)
    If Len(Trim(ValorICMSSubst.Text)) > 0 Then dValorICMSSubst = CDbl(ValorICMSSubst)
    If Len(Trim(ValorFrete.Text)) > 0 Then dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.Text)) > 0 Then dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(OutrasDespesas.Text)) > 0 Then dOutrasDespesas = CDbl(OutrasDespesas.Text)
    If Len(Trim(ValorIPI.Text)) > 0 Then dValorIPI = CDbl(ValorIPI.Text)

    dSoma = dValorProdutos + dValorICMSSubst + dValorFrete + dValorSeguro + dOutrasDespesas + dValorIPI

    If Format(dSoma, "0,00") <> Format(CDbl(ValorTotal.Text), "0,00") Then Error 17772

    'Move os dados da Tela para objCreditoPagar
    lErro = Move_Tela_Memoria(objCreditoPagar)
    If lErro <> SUCESSO Then Error 17202

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    If objCreditoPagar.dtDataEmissao <> DATA_NULA Then
        lErro = objContabil.Contabil_Testa_Data(objCreditoPagar.dtDataEmissao)
        If lErro <> SUCESSO Then Error 20829
    End If

    'Chama CreditoPagar_Grava (gravando inclusive os dados contábeis)
    lErro = CF("CreditoPagar_Grava", objCreditoPagar, objContabil)
    If lErro <> SUCESSO Then Error 17109

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 17103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 17105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case 17106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NAO_INFORMADO", Err)

        Case 17108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", Err)

        Case 17109, 17202, 20829 'Tratado na rotina chamada

        Case 17771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPRODUTOS_NAO_INFORMADO", Err)

        Case 17772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_INVALIDO", Err, ValorTotal.Text, dSoma)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155271)

    End Select

    Exit Function

End Function

Private Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorFrete_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorFrete.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
    If lErro <> SUCESSO Then Error 17765

    'Põe o valor formatado na tela
    ValorFrete.Text = Format(ValorFrete.Text, "Fixed")

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True


    Select Case Err

        Case 17765

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155272)

    End Select

    Exit Sub

End Sub

Private Sub ValorICMS_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorICMS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorICMS_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorICMS.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorICMS.Text)
    If lErro <> SUCESSO Then Error 17762

    'Põe o valor formatado na tela
    ValorICMS.Text = Format(ValorICMS.Text, "Fixed")

    Exit Sub

Erro_ValorICMS_Validate:

    Cancel = True


    Select Case Err

        Case 17762

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155273)

    End Select

    Exit Sub

End Sub

Private Sub ValorICMSSubst_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorICMSSubst_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorICMSSubst_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorICMSSubst.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorICMSSubst.Text)
    If lErro <> SUCESSO Then Error 17763

    'Põe o valor formatado na tela
    ValorICMSSubst.Text = Format(ValorICMSSubst.Text, "Fixed")

    Exit Sub

Erro_ValorICMSSubst_Validate:

    Cancel = True


    Select Case Err

        Case 17763

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155274)

    End Select

    Exit Sub

End Sub

Private Sub ValorIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorIPI_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIPI_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIPI.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorIPI.Text)
    If lErro <> SUCESSO Then Error 17766

    'Põe o valor formatado na tela
    ValorIPI.Text = Format(ValorIPI.Text, "Fixed")

    Exit Sub

Erro_ValorIPI_Validate:

    Cancel = True


    Select Case Err

        Case 17766

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155275)

    End Select

    Exit Sub

End Sub

Private Sub ValorIRRF_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorIRRF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIRRF_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIRRF.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
    If lErro <> SUCESSO Then Error 17764

    'Põe o valor formatado na tela
    ValorIRRF.Text = Format(ValorIRRF.Text, "Fixed")

    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True


    Select Case Err

        Case 17764

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155276)

    End Select

    Exit Sub

End Sub

Private Sub ValorProdutos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorProdutos_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorProdutos_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorProdutos.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(ValorProdutos.Text)
    If lErro <> SUCESSO Then Error 17769

    'Põe o valor formatado na tela
    ValorProdutos.Text = Format(ValorProdutos.Text, "Fixed")

    Exit Sub

Erro_ValorProdutos_Validate:

    Cancel = True


    Select Case Err

        Case 17769

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155277)

    End Select

    Exit Sub

End Sub

Private Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorSeguro.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
    If lErro <> SUCESSO Then Error 17767

    'Põe o valor formatado na tela
    ValorSeguro.Text = Format(ValorSeguro.Text, "Fixed")

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True


    Select Case Err

        Case 17767

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155278)

    End Select

    Exit Sub

End Sub

Private Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorTotal.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(ValorTotal.Text)
    If lErro <> SUCESSO Then Error 17770

    'Põe o valor formatado na tela
    ValorTotal.Text = Format(ValorTotal.Text, "Fixed")

    Exit Sub

Erro_ValorTotal_Validate:

    Cancel = True


    Select Case Err

        Case 17770

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155279)

    End Select

    Exit Sub

End Sub

'inicio contabilidade

Private Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Private Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Private Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

End Sub

Private Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Private Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub


Private Sub CTBGridContabil_LeaveCell()

    Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Private Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Private Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Private Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Private Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Private Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Private Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Private Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Private Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Private Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Private Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Private Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Private Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Private Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Private Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Private Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Private Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Private Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Private Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Private Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Private Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Private Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Private Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Private Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Private Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Private Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Private Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Private Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Private Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Private Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Private Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, sContaTela As String
Dim objFornecedor As New ClassFornecedor, objTipoFornecedor As New ClassTipoFornecedor
Dim objFilial As New ClassFilialFornecedor

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case VALORTOTAL1
            
            If Len(Trim(ValorTotal.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorTotal.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case FORNECEDOR_COD
            
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then Error 39603
                
                objMnemonicoValor.colValor.Add objFornecedor.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case FORNECEDOR_NOME
        
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then Error 39604
            
                objMnemonicoValor.colValor.Add objFornecedor.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case FILIAL_COD
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                objMnemonicoValor.colValor.Add objFilial.iCodFilial
            
            Else
                
                objMnemonicoValor.colValor.Add 0
            
            End If
            
        Case FILIAL_NOME_RED
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 39605
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 39606
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 41982
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 39607
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case NUMERO1
            
            If Len(Trim(NumTitulo.ClipText)) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NumTitulo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case EMISSAO1
            If Len(DataEmissao.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case VALOR_ICMS
            If Len(Trim(ValorICMS.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorICMS.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case VALOR_ICMS_SUBST
            If Len(Trim(ValorICMSSubst.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorICMSSubst.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case DEBITA_ICMS
            objMnemonicoValor.colValor.Add DebitoICMS.Value
            
        Case VALOR_IR
            If Len(Trim(ValorIRRF.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorIRRF.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case PIS_RETIDO
            If Len(Trim(PISRetido.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(PISRetido.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case COFINS_RETIDO
            If Len(Trim(COFINSRetido.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(COFINSRetido.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CSLL_RETIDO
            If Len(Trim(CSLLRetido.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(CSLLRetido.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_PRODUTOS
            If Len(Trim(ValorProdutos.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorProdutos.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case VALOR_FRETE
            If Len(Trim(ValorFrete.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorFrete.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_SEGURO
            If Len(Trim(ValorSeguro.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorSeguro.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_IPI
            If Len(Trim(ValorIPI.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorIPI.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case DEBITA_IPI
            objMnemonicoValor.colValor.Add DebitoIPI.Value
        
        Case OUTRAS_DESPESAS
            If Len(OutrasDespesas.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(OutrasDespesas.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
                   
        Case TIPO1
            If Len(Trim(Tipo.Text)) > 0 Then
                objMnemonicoValor.colValor.Add Tipo.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case CONTA_DESP_ESTOQUE
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then Error 41983
                
                objTipoFornecedor.iCodigo = objFornecedor.iTipo
                lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
                If lErro <> SUCESSO Then Error 41984
                
                If objTipoFornecedor.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTipoFornecedor.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then Error 41985
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case Else
            Error 39608
            
    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 39608
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 39603, 39604, 39605, 39606, 39607, 41982 To 41985
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155280)

    End Select

    Exit Function

End Function

Private Sub LabelTipo_Click()

Dim objTipoDocumento As New ClassTipoDocumento
Dim colSelecao As Collection

    objTipoDocumento.sSigla = Tipo.Text
    
    'Chama a tela TipoDocOutrosPagLista
    Call Chama_Tela("TipoDocCreditoPagLista", colSelecao, objTipoDocumento, objEventoTipo)
    
End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDocumento = obj1

    'Preenche campo Tipo
    Tipo.Text = objTipoDocumento.sSigla
    
    Call Tipo_Validate(bSGECancelDummy)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTipo_evSelecao:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155281)
     
     End Select
     
     Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_DEVOL_CREDITOS_ID
    Set Form_Load_Ocx = Me
    Caption = "Devoluções / Créditos com Fornecedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CreditosPagar"
    
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
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call LabelTipo_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        End If
    
    End If
    
End Sub



Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
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

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub PISRetido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PISRetido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISRetido_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(PISRetido.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(PISRetido.Text)
    If lErro <> SUCESSO Then Error 17764

    'Põe o valor formatado na tela
    PISRetido.Text = Format(PISRetido.Text, "Fixed")

    Exit Sub

Erro_PISRetido_Validate:

    Cancel = True


    Select Case Err

        Case 17764

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155282)

    End Select

    Exit Sub

End Sub

Private Sub COFINSRetido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub COFINSRetido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSRetido_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(COFINSRetido.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(COFINSRetido.Text)
    If lErro <> SUCESSO Then Error 17764

    'Põe o valor formatado na tela
    COFINSRetido.Text = Format(COFINSRetido.Text, "Fixed")

    Exit Sub

Erro_COFINSRetido_Validate:

    Cancel = True


    Select Case Err

        Case 17764

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155283)

    End Select

    Exit Sub

End Sub

Private Sub CSLLRetido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CSLLRetido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CSLLRetido_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(CSLLRetido.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(CSLLRetido.Text)
    If lErro <> SUCESSO Then Error 17764

    'Põe o valor formatado na tela
    CSLLRetido.Text = Format(CSLLRetido.Text, "Fixed")

    Exit Sub

Erro_CSLLRetido_Validate:

    Cancel = True


    Select Case Err

        Case 17764

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155284)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Preenche()

Static sNomeReduzidoParte As String
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134048

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134048

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155285)

    End Select
    
    Exit Sub

End Sub
Private Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Private Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub

Private Sub BotaoBaixas_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoBaixas_Click

    'Verifica se o Sequencial foi informado
    If glNumIntDoc = 0 Then gError 15451
    
    'Filtro
    colSelecao.Add MOTIVO_CREDITO_FORNECEDOR
    colSelecao.Add glNumIntDoc

    'Abre o Browse de Antecipações de recebimento de uma Filial
    Call Chama_Tela("BaixasPagLista", colSelecao, Nothing, Nothing, "NumIntBaixa IN (SELECT NumIntBaixa FROM BaixasPag WHERE Motivo = ? AND NumIntDoc = ? AND Status <> 5)")

    Exit Sub

Erro_BotaoBaixas_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            
        Case 15451
            Call Rotina_Erro(vbOKOnly, "ERRO_ANTECIPRECEB_NAO_CARREGADO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142882)

    End Select

    Exit Sub

End Sub

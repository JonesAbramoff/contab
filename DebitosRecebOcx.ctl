VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl DebitosRecebOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5010
      Index           =   1
      Left            =   225
      TabIndex        =   0
      Top             =   795
      Width           =   9075
      Begin VB.Frame Frame6 
         Caption         =   "Situação Atual"
         Height          =   765
         Left            =   255
         TabIndex        =   103
         Top             =   4260
         Width           =   8415
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
            TabIndex        =   104
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label Saldo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   285
            Left            =   1725
            TabIndex        =   106
            Top             =   308
            Width           =   1530
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
            TabIndex        =   105
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   "Dados Principais"
         Height          =   1935
         Left            =   240
         TabIndex        =   48
         Top             =   -30
         Width           =   8400
         Begin VB.TextBox Observacao 
            Height          =   300
            Left            =   1530
            MaxLength       =   255
            TabIndex        =   7
            Top             =   1515
            Width           =   6135
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5835
            TabIndex        =   2
            Top             =   330
            Width           =   1815
         End
         Begin VB.ComboBox Tipo 
            Height          =   315
            Left            =   1530
            TabIndex        =   3
            Text            =   " "
            Top             =   720
            Width           =   2775
         End
         Begin MSMask.MaskEdBox NumTitulo 
            Height          =   300
            Left            =   5835
            TabIndex        =   4
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1530
            TabIndex        =   1
            Top             =   330
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
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1125
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
            Top             =   1125
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
            Left            =   5835
            TabIndex        =   6
            Top             =   1125
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
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
            Left            =   345
            TabIndex        =   102
            Top             =   1530
            Width           =   1080
         End
         Begin VB.Label LabelFilial 
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
            Top             =   390
            Width           =   525
         End
         Begin VB.Label ClienteLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
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
            Left            =   780
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   56
            Top             =   390
            Width           =   660
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
            TabIndex        =   57
            Top             =   765
            Width           =   720
         End
         Begin VB.Label Label3 
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
            TabIndex        =   58
            Top             =   1185
            Width           =   765
         End
         Begin VB.Label Label7 
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
            TabIndex        =   59
            Top             =   1185
            Width           =   510
         End
         Begin VB.Label TipoLabel 
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
            TabIndex        =   60
            Top             =   780
            Width           =   450
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Valores"
         Height          =   2265
         Left            =   255
         TabIndex        =   49
         Top             =   1965
         Width           =   8400
         Begin VB.Frame Frame2 
            Caption         =   "Retenções"
            Height          =   1035
            Left            =   3090
            TabIndex        =   92
            Top             =   1125
            Width           =   4920
            Begin MSMask.MaskEdBox ValorIRRF 
               Height          =   300
               Left            =   1110
               TabIndex        =   93
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
               TabIndex        =   94
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
               TabIndex        =   95
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
               TabIndex        =   96
               Top             =   645
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
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
               TabIndex        =   100
               Top             =   285
               Width           =   300
            End
            Begin VB.Label Label4 
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
               TabIndex        =   99
               Top             =   315
               Width           =   375
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
               TabIndex        =   98
               Top             =   660
               Width           =   735
            End
            Begin VB.Label Label1 
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
               TabIndex        =   97
               Top             =   690
               Width           =   525
            End
         End
         Begin MSMask.MaskEdBox ValorProdutos 
            Height          =   300
            Left            =   6930
            TabIndex        =   10
            Top             =   270
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   300
            Left            =   885
            TabIndex        =   11
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox OutrasDespesas 
            Height          =   300
            Left            =   6945
            TabIndex        =   14
            Top             =   735
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   300
            Left            =   3675
            TabIndex        =   12
            Top             =   705
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorIPI 
            Height          =   300
            Left            =   885
            TabIndex        =   13
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorICMS 
            Height          =   300
            Left            =   870
            TabIndex        =   8
            Top             =   240
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
            Left            =   3675
            TabIndex        =   9
            Top             =   270
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   6045
            TabIndex        =   61
            Top             =   330
            Width           =   825
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   285
            TabIndex        =   62
            Top             =   780
            Width           =   510
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   5355
            TabIndex        =   63
            Top             =   795
            Width           =   1515
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   2955
            TabIndex        =   64
            Top             =   765
            Width           =   675
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   465
            TabIndex        =   65
            Top             =   1260
            Width           =   315
         End
         Begin VB.Label Label14 
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
            Left            =   270
            TabIndex        =   66
            Top             =   285
            Width           =   525
         End
         Begin VB.Label Label16 
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
            Left            =   2535
            TabIndex        =   67
            Top             =   330
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4605
      Index           =   3
      Left            =   225
      TabIndex        =   22
      Top             =   885
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4920
         TabIndex        =   101
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
         Left            =   7755
         TabIndex        =   27
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   870
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
         TabIndex        =   26
         Top             =   30
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
         Left            =   6300
         TabIndex        =   28
         Top             =   345
         Width           =   2700
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   38
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   37
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   40
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   50
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
            Height          =   195
            Left            =   240
            TabIndex        =   68
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
            Height          =   195
            Left            =   1125
            TabIndex        =   69
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   70
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   71
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
         TabIndex        =   31
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4920
         TabIndex        =   36
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
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   32
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
         Left            =   3435
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   51
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   39
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
         TabIndex        =   41
         Top             =   1515
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
         TabIndex        =   42
         Top             =   1515
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
         Left            =   6360
         TabIndex        =   29
         Top             =   660
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
         Height          =   255
         Left            =   45
         TabIndex        =   72
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   73
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
         TabIndex        =   74
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   75
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   76
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
         TabIndex        =   77
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
         TabIndex        =   78
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
         TabIndex        =   79
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
         TabIndex        =   80
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
         TabIndex        =   81
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
         Height          =   225
         Left            =   1800
         TabIndex        =   82
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   83
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   84
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
         TabIndex        =   85
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
         Height          =   195
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   86
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
         Height          =   195
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   87
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6165
      ScaleHeight     =   495
      ScaleWidth      =   3135
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   45
      Width           =   3195
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
         Picture         =   "DebitosRecebOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   390
         Left            =   1125
         Picture         =   "DebitosRecebOcx.ctx":0F0A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   390
         Left            =   1635
         Picture         =   "DebitosRecebOcx.ctx":1064
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   390
         Left            =   2145
         Picture         =   "DebitosRecebOcx.ctx":11EE
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   390
         Left            =   2655
         Picture         =   "DebitosRecebOcx.ctx":1720
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4545
      Index           =   2
      Left            =   255
      TabIndex        =   15
      Top             =   870
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame Frame3 
         Caption         =   "Comissões"
         Height          =   4050
         Left            =   120
         TabIndex        =   52
         Top             =   255
         Width           =   8400
         Begin VB.CommandButton Vendedores 
            Caption         =   "Vendedores"
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
            Left            =   7020
            Picture         =   "DebitosRecebOcx.ctx":189E
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   3315
            Width           =   1215
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   285
            Left            =   1095
            TabIndex        =   16
            Top             =   165
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   225
            Left            =   5580
            TabIndex        =   19
            Top             =   165
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   225
            Left            =   4335
            TabIndex        =   18
            Top             =   165
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
         Begin MSMask.MaskEdBox Percentual 
            Height          =   280
            Left            =   3210
            TabIndex        =   17
            Top             =   150
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   1155
            Left            =   420
            TabIndex        =   20
            Top             =   315
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   2037
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label TotalPercentual 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1920
            TabIndex        =   88
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label TotalValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3150
            TabIndex        =   89
            Top             =   1545
            Width           =   1155
         End
         Begin VB.Label TotalLabel 
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
            Height          =   225
            Left            =   1260
            TabIndex        =   90
            Top             =   1545
            Width           =   705
         End
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5565
      Left            =   120
      TabIndex        =   53
      Top             =   345
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9816
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "DebitosRecebOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Private iClienteAlterado As Integer
Private iFrameAtual As Integer
Private iValorAlterado As Integer
Private iSubTipoAtual As Integer
Private glNumIntDoc As Long

Dim objGridComissoes As AdmGrid
Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Debito_Col As Integer
Dim iGrid_Credito_Col As Integer
Dim iGrid_Historico_Col As Integer
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_Percentual_Col As Integer
Dim iGrid_ValorBase_Col As Integer
Dim iGrid_Valor_Col As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoTipoDocumento As AdmEvento
Attribute objEventoTipoDocumento.VB_VarHelpID = -1

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'Mnemônicos
Private Const CLIENTE_COD As String = "Cliente_Codigo"
Private Const CLIENTE_NOME As String = "Cliente_Nome"
Private Const FILIAL_COD As String = "FilialCli_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialCli_Nome"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialCli_CGC_CPF"
Private Const TIPO1 As String = "Tipo"
Private Const NUM_TITULO As String = "Numero_Titulo"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const VALOR_TOTAL As String = "Valor_Total"
Private Const VALOR_ICMS As String = "Valor_ICMS"
Private Const VALOR_ICMS_SUBST As String = "Valor_ICMS_Subst"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_IPI As String = "Valor_IPI"
Private Const OUTRAS_DESPESAS As String = "Outras_Despesas"
Private Const CONTA_TIPO_CLIENTE As String = "Conta_Tipo_Cliente"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Comissoes = 2
Private Const TAB_Contabilizacao = 3

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim lNumIntNF As Long
Dim objCliente As New ClassCliente
Dim objDebitoReceber As New ClassDebitoRecCli
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoDocOriginal_Click
    
    'Critica se os Campos estão Preenchidos(Fornecedor, Filial, Tipo, Número, DataEmissao)
    lErro = DocOriginal_Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 79649

    objCliente.sNomeReduzido = Cliente.Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79650

    'Não achou o Cliente --> erro
    If lErro <> SUCESSO Then gError 79651
    
    'Preenche objCreditoPagar
    objDebitoReceber.lCliente = objCliente.lCodigo
    objDebitoReceber.iFilial = Codigo_Extrai(Filial.Text)
    objDebitoReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objDebitoReceber.lNumTitulo = CLng(NumTitulo.Text)
    objDebitoReceber.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objDebitoReceber.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objDebitoReceber.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Débito (Baixados ou não)
    lErro = CF("DebitoRecCli_Le_Numero", objDebitoReceber)
    If lErro <> SUCESSO And lErro <> 17916 And lErro <> 17917 Then gError 79652
    
    'Se não encontrou o título => erro
    If lErro = 17916 Then gError 79653

    'Se o documento é do tipo NCR ou NER => erro (não possui doc original)
    If objDebitoReceber.sSiglaDocumento = SIGLA_NOTA_DEBITO_RECEBER Or objDebitoReceber.sSiglaDocumento = SIGLA_NOTA_ENTRADA_RETORNO Then gError 79654
    
    'Verifica se o Débito foi gerado através do CR ou através de uma NFiscal do EST
    lErro = CF("DocumentoCPR_OrigemNFiscal", objDebitoReceber.sSiglaDocumento, objDebitoReceber.lNumIntDoc, lNumIntNF)
    If lErro <> SUCESSO And lErro <> 41542 Then gError 79655
    
    'Se não encontrou NFiscal de origem => erro
    If lErro = 41542 Then gError 79656
    
    objNFiscal.lNumIntDoc = lNumIntNF
        
    'Se a versão utilizada for a Full
    If giTipoVersao = VERSAO_FULL Then
        'Chama a Tela 'Entrada - Nota Fiscal de Devolução'
        Call Chama_Tela("NFiscalEntDev", objNFiscal)
    
    'Se a versão utilizada for a Light
    ElseIf giTipoVersao = VERSAO_LIGHT Then
        
        'Chama a Tela 'Notas Fiscais de Entrada'
        Call Chama_Tela("NFiscalFatEntrada", objNFiscal)
    End If
    
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 79649, 79650, 79652, 79655
        
        Case 79651
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case 79654, 79656
           Call Rotina_Erro(vbOKOnly, "ERRO_DEBITO_RECEBER_SEM_DOC_ORIGINAL", gErr, objDebitoReceber.lNumTitulo)
        
        Case 79653
            
            'Se a data não estiver preenchida
            If objDebitoReceber.dtDataEmissao = DATA_NULA Then
                'Não exibe o parâmetro data
                Call Rotina_Erro(vbOKOnly, "ERRO_DEBITORECCLI_NAO_CADASTRADO", gErr, objDebitoReceber.lCliente, objDebitoReceber.iFilial, objDebitoReceber.sSiglaDocumento, objDebitoReceber.lNumTitulo, "")
            'Senão
            Else
                'O parâmetro data será exibido
                Call Rotina_Erro(vbOKOnly, "ERRO_DEBITORECCLI_NAO_CADASTRADO", gErr, objDebitoReceber.lCliente, objDebitoReceber.iFilial, objDebitoReceber.sSiglaDocumento, objDebitoReceber.lNumTitulo, objDebitoReceber.dtDataEmissao)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158778)

    End Select

    Exit Sub

End Sub

Private Function DocOriginal_Critica_CamposPreenchidos() As Long

On Error GoTo Erro_DocOriginal_Critica_CamposPreenchidos

    'Se o Cliente não foi preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 79644
    
    'Se a Filial não foi selecionada => erro
    If Len(Trim(Filial.Text)) = 0 Then gError 79645
    
    'Se o tipo do Documento não foi selecionado => erro
    If Len(Trim(Tipo.Text)) = 0 Then gError 79646
    
    'Se o número do Documento não foi preenchido => erro
    If Len(Trim(NumTitulo.Text)) = 0 Then gError 79647
    
    DocOriginal_Critica_CamposPreenchidos = SUCESSO
    
    Exit Function
    
Erro_DocOriginal_Critica_CamposPreenchidos:

    DocOriginal_Critica_CamposPreenchidos = gErr
    
    Select Case gErr
    
        Case 79644
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 79645
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79646
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
        Case 79647
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158779)
        
    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos obrigatórios estão preenchidos
    If Len(Trim(Cliente.Text)) = 0 Then Error 17908
    If Len(Trim(Filial.Text)) = 0 Then Error 17909
    If Len(Trim(Tipo.Text)) = 0 Then Error 17910
    If Len(Trim(NumTitulo.Text)) = 0 Then Error 17911

    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objDebitoRecCli.dtDataEmissao = DATA_NULA
    Else
        objDebitoRecCli.dtDataEmissao = CDate(DataEmissao.Text)
    End If

    lErro = Move_Tela_Memoria(objDebitoRecCli, colInfoComissao)
    If lErro <> SUCESSO Then Error 17912

    'Verifica se campos identificadores correspondem a um Débito a Receber Cliente
    lErro = CF("DebitoRecCli_Le_Numero", objDebitoRecCli)
    If lErro <> SUCESSO And lErro <> 17916 And lErro <> 17917 Then Error 17918

    If lErro = 17916 Then Error 17331

    If lErro = 17917 Then Error 17332

    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_DEBITORECCLI", objDebitoRecCli.lNumTitulo)

    'Se não confirmar, sai
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui Débito a Receber Cliente (inclusive os dados contábeis)(contabilidade)
    lErro = CF("DebitoRecCli_Exclui", objDebitoRecCli, objContabil)
    If lErro <> SUCESSO Then Error 17919

    'Limpa a Tela
    Call Limpa_Tela_DebitoRecCli

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DEBITORECCLI_NAO_ENCONTRADO1", Err, objDebitoRecCli.lCliente, objDebitoRecCli.iFilial, objDebitoRecCli.sSiglaDocumento, objDebitoRecCli.lNumTitulo)

        Case 17332
            lErro = Rotina_Erro(vbOKOnly, "AVISO_NAO_E_PERMITIDO_EXCLUSAO_DEBRECCLI_BAIXADO", Err, objDebitoRecCli.lCliente, objDebitoRecCli.iFilial, objDebitoRecCli.sSiglaDocumento, objDebitoRecCli.lNumTitulo)

        Case 17908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 17909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 17910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", Err)

        Case 17911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case 17912, 17918, 17919

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158780)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 17871

    'Limpa a Tela
    Call Limpa_Tela_DebitoRecCli

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 17871

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158781)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17903

    'Limpa a Tela
    Call Limpa_Tela_DebitoRecCli

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 17903

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158782)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = 1

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        If Len(Trim(Cliente.Text)) > 0 Then

            lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then Error 17803

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then Error 17804

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            If iCodFilial = FILIAL_MATRIZ Then
                Filial.ListIndex = 0
            Else
                Call CF("Filial_Seleciona", Filial, iCodFilial)
            End If

        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            Filial.Clear

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case Err

        Case 17803
            
        Case 17804

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158783)

    End Select

    Exit Sub

End Sub

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Cliente.Text)) > 0 Then objCliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

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

    'Verifica se a data de emissão está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Verifica se a data final é válida
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 17814

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 17814

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158784)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO
       
    If Filial.ListIndex = -1 Then Exit Sub
    
    'Verifica se Cliente, Filial e Valor estão preenchidos e se o Grid de Comissões Emissão está vazio
    If Len(Trim(Cliente.ClipText)) > 0 And Len(Trim(Filial.Text)) > 0 And Len(Trim(ValorTotal.ClipText)) > 0 And objGridComissoes.iLinhasExistentes = 0 Then
    
        lErro = Inicializa_Comissao()
        If lErro <> SUCESSO Then Error 43510
        
    End If
    
    Exit Sub
    
Erro_Filial_Click:

    Select Case Err
    
        Case 43510
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158785)
            
    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 17805

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then Error 17806

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 17807

        If lErro = 17660 Then Error 17808

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        'Verifica se Cliente, Filial e Valor estão preenchidos e se o Grid de Comissões Emissão está vazio
        If Len(Trim(Cliente.ClipText)) > 0 And Len(Trim(Valor.ClipText)) > 0 And objGridComissoes.iLinhasExistentes = 0 Then
        
            lErro = Inicializa_Comissao()
            If lErro <> SUCESSO Then Error 43511
            
        End If

    End If

    'Não encontrou a STRING
    If lErro = 6731 Then Error 17809

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

       Case 17805, 17807

       Case 17806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

       Case 17808
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 17809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)
        
        Case 43511
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158786)

    End Select

    Exit Sub

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
    Set objEventoCliente = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoTipoDocumento = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    Set objGridComissoes = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Function Inicializa_Grid_DebitosRecebComissoes(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_DebitosRecebComissoes

    'Tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Vendedor")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Valor Base")
    objGridInt.colColuna.Add ("Valor")

    'Campos de edição do grid
    objGridInt.colCampo.Add (Vendedor.Name)
    objGridInt.colCampo.Add (Percentual.Name)
    objGridInt.colCampo.Add (ValorBase.Name)
    objGridInt.colCampo.Add (Valor.Name)

    iGrid_Vendedor_Col = 1
    iGrid_Percentual_Col = 2
    iGrid_ValorBase_Col = 3
    iGrid_Valor_Col = 4

    objGridInt.objGrid = GridComissoes

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    GridComissoes.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalPercentual.top = GridComissoes.top + GridComissoes.Height
    TotalPercentual.left = GridComissoes.left
    For iIndice = 0 To 1
        TotalPercentual.left = TotalPercentual.left + GridComissoes.ColWidth(iIndice) + GridComissoes.GridLineWidth + 20
    Next

    TotalPercentual.Width = GridComissoes.ColWidth(2)

    TotalValor.top = TotalPercentual.top
    TotalValor.Width = GridComissoes.ColWidth(4)
    For iIndice = 0 To 3
        TotalValor.left = TotalPercentual.left + TotalPercentual.Width + GridComissoes.ColWidth(iIndice) + GridComissoes.GridLineWidth + 20
    Next

    TotalLabel.top = TotalPercentual.top + (TotalPercentual.Height - TotalLabel.Height) / 2
    TotalLabel.left = TotalPercentual.left - TotalLabel.Width

    Inicializa_Grid_DebitosRecebComissoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_DebitosRecebComissoes:

    Inicializa_Grid_DebitosRecebComissoes = Err

    Select Case Err

        Case 14251

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158787)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim colTipoDocumento As New colTipoDocumento
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iSubTipoAtual = 0
    
    'Visibilidade para versão LIGHT
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        LabelFilial.left = POSICAO_FORA_TELA
'        Filial.left = POSICAO_FORA_TELA
'        Filial.TabStop = False
'
'    End If
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
    Set objGridComissoes = New AdmGrid

    Set objEventoNumero = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoTipoDocumento = New AdmEvento
    
    lErro = Inicializa_Grid_DebitosRecebComissoes(objGridComissoes)
    If lErro <> SUCESSO Then Error 14250

    'Preenche a ComboBox com  os Tipos de Documentos existentes no BD
    lErro = CF("TiposDocumento_Le_DebReceber", colTipoDocumento)
    If lErro <> SUCESSO Then Error 17826

    For Each objTipoDocumento In colTipoDocumento

        'Preenche a ComboBox Tipo com os objetos da colecao colTipoDocumento
        Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida

    Next
    
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then Error 39679

    iAlterado = 0
    iValorAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 14250, 17826, 39679

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158788)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click

    'Verifica se campos Cliente, Filial, Tipo estão preenchidos
    If Len(Trim(Cliente.Text)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(Tipo.Text)) = 0 Then Error 17945

    lErro = Move_Tela_Memoria(objDebitoRecCli, colInfoComissao)
    If lErro <> SUCESSO Then Error 17946

    'Armazena fitro Cliente, Filial, Tipo para Browse
    colSelecao.Add objDebitoRecCli.lCliente
    colSelecao.Add objDebitoRecCli.iFilial
    colSelecao.Add objDebitoRecCli.sSiglaDocumento

    Call Chama_Tela("DebitosRecebLista", colSelecao, objDebitoRecCli, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 17945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_DEBITO_RECEBER_NAO_PREENCHIDOS", Err)

        Case 17946

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158789)

    End Select

    Exit Sub

End Sub

Private Sub NumTitulo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumTitulo_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)
    
End Sub

Private Sub NumTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumTitulo_Validate

    If Len(Trim(NumTitulo.ClipText)) > 0 Then

        If Not IsNumeric(NumTitulo.ClipText) Then Error 17812

        If CLng(NumTitulo) < 1 Then Error 17813

    End If

    Exit Sub

Erro_NumTitulo_Validate:

    Cancel = True


    Select Case Err

        Case 17812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_E_NUMERICO", Err, NumTitulo.Text)

        Case 17813
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", Err, NumTitulo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158790)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente, Cancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(Cancel)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDebitoRecCli As ClassDebitoRecCli

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objDebitoRecCli = obj1

    lErro = Traz_Debito_Tela(objDebitoRecCli)
    If lErro <> SUCESSO Then Error 17947

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 17947

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158791)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_objEventoVendedor_evSelecao

    If GridComissoes.Row >= GridComissoes.FixedRows Then

        Set objVendedor = obj1
        
        Call Vendedor_Linha_Preenche(objVendedor)
        
    End If

    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158792)

        End Select

    Exit Sub

End Sub

Private Sub objEventoTipoDocumento_evSelecao(obj1 As Object)

Dim objTipoDoc As ClassTipoDocumento
Dim lErro As Long

On Error GoTo Erro_objEventoTipoDocumento_evSelecao

    Set objTipoDoc = obj1

    'Preenche campo Tipo
    Tipo.Text = objTipoDoc.sSigla

    'Executa o Validate
    Call Tipo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoTipoDocumento_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158793)

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
                Parent.HelpContextID = IDH_DEVOL_DEB_CLIENTES_ID
                
            Case TAB_Comissoes
                Parent.HelpContextID = IDH_DEVOL_DEB_CLIENTES_COMISSSOES
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_DEVOL_DEB_CLIENTES_CONTABILIZACAO
                        
        End Select

    End If

End Sub

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
    If lErro <> SUCESSO Then Error 17823

    'Põe o valor formatado na tela
    OutrasDespesas.Text = Format(OutrasDespesas.Text, "Fixed")

    Exit Sub

Erro_OutrasDespesas_Validate:

    Cancel = True


    Select Case Err

        Case 17823

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158794)

    End Select

    Exit Sub

End Sub

Private Sub Percentual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Click()

Dim lErro As Long, iSubTipo As Integer

On Error GoTo Error_Tipo_Click

    iAlterado = REGISTRO_ALTERADO

    'Processa alteração do Subtipo para que os novos modelos sejam carregados
    Select Case SCodigo_Extrai(Tipo.Text)
        Case "DCLI"
            iSubTipo = 2
        Case Else
            iSubTipo = 0
    End Select
    
    If iSubTipoAtual <> iSubTipo Then
    
        lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(iSubTipo)
        If lErro <> SUCESSO Then Error 61833
        
        iSubTipoAtual = iSubTipo
        
    End If
    
    Exit Sub

Error_Tipo_Click:

    Select Case Err

        Case 61833
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158795)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Error_Tipo_Validate

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Tipo
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = CF("SCombo_Seleciona", Tipo)
    If lErro <> SUCESSO And lErro <> 60483 Then Error 61833

    'Se nao encontrar -> Erro
    If lErro = 60483 Then Error 17811

    Exit Sub

Error_Tipo_Validate:

    Cancel = True

    Select Case Err

        Case 17811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_NAO_ENCONTRADO", Err, Tipo.Text)

        Case 61833
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158795)

    End Select

    Exit Sub

End Sub

Private Sub TipoLabel_Click()

Dim objTipoDoc As New ClassTipoDocumento
Dim colSelecao As New Collection

    If Len(Tipo.Text) > 0 Then
        objTipoDoc.sSigla = Tipo.Text
    Else
        objTipoDoc.sSigla = ""
    End If

    Call Chama_Tela("TipoDocDebitosRecLista", colSelecao, objTipoDoc, objEventoTipoDocumento)

    Exit Sub

End Sub

Private Sub UpDownEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    DataEmissao.SetFocus

    If Len(DataEmissao.ClipText) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 17815

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 17815

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158796)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    DataEmissao.SetFocus

    If Len(DataEmissao.ClipText) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 17816

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 17816

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158797)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

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
    If lErro <> SUCESSO Then Error 17821

    'Põe o valor formatado na tela
    ValorFrete.Text = Format(ValorFrete.Text, "Fixed")

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True


    Select Case Err

        Case 17821

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158798)

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
    If lErro <> SUCESSO Then Error 17818

    'Põe o valor formatado na tela
    ValorICMS.Text = Format(ValorICMS.Text, "Fixed")

    Exit Sub

Erro_ValorICMS_Validate:

    Cancel = True


    Select Case Err

        Case 17818

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158799)

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

    'critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorICMSSubst.Text)
    If lErro <> SUCESSO Then Error 17819

    'Põe o valor formatado na tela
    ValorICMSSubst.Text = Format(ValorICMSSubst.Text, "Fixed")

    Exit Sub

Erro_ValorICMSSubst_Validate:

    Cancel = True


    Select Case Err

        Case 17819

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158800)

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

    'critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorIPI.Text)
    If lErro <> SUCESSO Then Error 17824

    'Põe o valor formatado na tela
    ValorIPI.Text = Format(ValorIPI.Text, "Fixed")

    Exit Sub

Erro_ValorIPI_Validate:

    Cancel = True


    Select Case Err

        Case 17824

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158801)

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

    'critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
    If lErro <> SUCESSO Then Error 17825

    'Põe o valor formatado na tela
    ValorIRRF.Text = Format(ValorIRRF.Text, "Fixed")

    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True


    Select Case Err

        Case 17825

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158802)

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
    If lErro <> SUCESSO Then Error 17820

    'Põe o valor formatado na tela
    ValorProdutos.Text = Format(ValorProdutos.Text, "Fixed")

    Exit Sub

Erro_ValorProdutos_Validate:

    Cancel = True


    Select Case Err

        Case 17820

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158803)

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
    If lErro <> SUCESSO Then Error 17822

    'Põe o valor formatado na tela
    ValorSeguro.Text = Format(ValorSeguro.Text, "Fixed")

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True


    Select Case Err

        Case 17822

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158804)

    End Select

    Exit Sub

End Sub

Private Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Validate

    'Verifica se o Valor foi Alterado
    If iValorAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(ValorTotal.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(ValorTotal.Text)
    If lErro <> SUCESSO Then Error 17817

    'Põe o valor formatado na tela
    ValorTotal.Text = Format(ValorTotal.Text, "Fixed")

    'Verifica se Cliente, Filial e Valor estão preenchidos
    If Len(Trim(Cliente.ClipText)) > 0 And Len(Trim(Filial.Text)) > 0 Then
    
        lErro = Inicializa_Comissao()
        If lErro <> SUCESSO Then Error 43509
        
    End If

    iValorAlterado = 0
    
    Exit Sub

Erro_ValorTotal_Validate:

    Cancel = True


    Select Case Err

        Case 17817
            
        Case 43509

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158805)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Comissao() As Long
'Inicializa as Comissões para o vendedor default da filial do cliente ou do tipo de cliente

Dim lErro As Long, iCodFilial As Integer
Dim objVendedor As New ClassVendedor
Dim objComissao As New ClassComissao

On Error GoTo Erro_Inicializa_Comissao

    If Len(Trim(ValorTotal.Text)) <> 0 Then
    
        iCodFilial = Codigo_Extrai(Filial.Text)
        lErro = ComissaoAutomatica_Obter_Debito(Cliente.Text, iCodFilial, objVendedor, objComissao, StrParaDbl(ValorTotal.Text), StrParaDbl(ValorFrete.Text), StrParaDbl(ValorSeguro.Text), StrParaDbl(OutrasDespesas.Text), StrParaDbl(ValorIPI.Text), StrParaDbl(ValorICMSSubst.Text))
        If lErro <> SUCESSO Then Error 59209
        
        'Limpa o Grid
        Call Grid_Limpa(objGridComissoes)

        'Mostra na tela o Vendedor,Percentual Comissão, Valor Base e Valor Comissão
        If objVendedor.sNomeReduzido <> "" Then
            
            GridComissoes.TextMatrix(1, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
            GridComissoes.TextMatrix(1, iGrid_Percentual_Col) = Format(objComissao.dPercentual, "Percent")
            GridComissoes.TextMatrix(1, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
            GridComissoes.TextMatrix(1, iGrid_Valor_Col) = Format(objComissao.dValor, "Standard")
            
            objGridComissoes.iLinhasExistentes = 1
        
        End If
        
        'Chama Soma_Percentual
        Call Soma_Percentual
        
        'Chama Soma_Valor
        Call Soma_Valor
    
    End If
    
    Inicializa_Comissao = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Comissao:

    Inicializa_Comissao = Err
    
    Select Case Err
    
        Case 59209
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158806)
            
    End Select
        
    Exit Function
    
End Function

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Percentual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub Percentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub Percentual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Percentual
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridComissoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Private Sub GridComissoes_GotFocus()

    Call Grid_Recebe_Foco(objGridComissoes)

End Sub

Private Sub GridComissoes_EnterCell()

    Call Grid_Entrada_Celula(objGridComissoes, iAlterado)

End Sub

Private Sub GridComissoes_LeaveCell()

    Call Saida_Celula(objGridComissoes)

End Sub

Private Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridComissoes)
    
    Call Soma_Percentual
    
    Call Soma_Valor

End Sub

Private Sub GridComissoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Private Sub GridComissoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComissoes)

End Sub

Private Sub GridComissoes_RowColChange()

    Call Grid_RowColChange(objGridComissoes)

End Sub

Private Sub GridComissoes_Scroll()

    Call Grid_Scroll(objGridComissoes)

End Sub

Function Trata_Parametros(Optional objDebitoRecCli As ClassDebitoRecCli) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um tipo de documento selecionado, exibir seus dados
    If Not (objDebitoRecCli Is Nothing) Then

        'Verifica se o debito a receber existe
        lErro = CF("DebitoReceber_Le", objDebitoRecCli)
        If lErro <> SUCESSO And lErro <> 17835 Then Error 17836

        If lErro = 17835 Then Error 17837

        lErro = Traz_Debito_Tela(objDebitoRecCli)
        If lErro <> SUCESSO Then Error 17838

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 17836, 17838

        Case 17837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DEBITORECCLI_NAO_ENCONTRADO", Err, objDebitoRecCli.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158807)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer
'Dim ColRateioOn As New Collection

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39681

        If objGridInt.objGrid Is GridComissoes Then

            Select Case objGridInt.objGrid.Col

                Case iGrid_Vendedor_Col
                    lErro = Saida_Celula_Vendedor(objGridInt)
                    If lErro <> SUCESSO Then Error 17855

                Case iGrid_Percentual_Col
                    lErro = Saida_Celula_Percentual(objGridInt)
                    If lErro <> SUCESSO Then Error 17856

                Case iGrid_ValorBase_Col
                    lErro = Saida_Celula_ValorBase(objGridInt)
                    If lErro <> SUCESSO Then Error 17857

                Case iGrid_Valor_Col
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then Error 17858

            End Select

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 17859

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 17855, 17856, 17857, 17858

        Case 17859
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 39681
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Percentual(objGridInt As AdmGrid) As Long
'Faz a crítica da celula PercentualComissoes do grid que está deixando de ser o corrente

Dim lErro As Long
Dim cPercentual As Currency
Dim dValorBase As Double
Dim dValorComissao As Double

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = Percentual

    'Verifica se o percentual está preenchido
    If Len(Trim(Percentual.ClipText)) > 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(Percentual.Text)
        If lErro <> SUCESSO Then Error 17839

        'Mostra na tela o percentual formatado
        Percentual.Text = Format(Percentual.Text, "Fixed")

        cPercentual = CDbl(Percentual.Text)

        'Verifica se valor base correspondente esta preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))

           'Calcula o valor comissão
           dValorComissao = (cPercentual * dValorBase) / 100

           'Coloca o valorcomissoes na tela
           GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Col) = Format(dValorComissao, "Standard")

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 17863

    'Chama Soma_Percentual
    lErro = Soma_Percentual()
    If lErro <> SUCESSO Then Error 17862

    'Chama Soma_Valor
    lErro = Soma_Valor()
    If lErro <> SUCESSO Then Error 17948

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = Err

    Select Case Err

        Case 17839, 17863
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 17862, 17948

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158808)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Vendedor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim objVendedor As New ClassVendedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Vendedor

    Set objGridInt.objControle = Vendedor

    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then

        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Grid(Vendedor, objVendedor)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then Error 17840

        If lErro = 25018 Then Error 17841

        If lErro = 25020 Then Error 17842
        
        lErro = Vendedor_Linha_Preenche(objVendedor)
        If lErro <> SUCESSO Then Error 17860
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 17861

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = Err

    Select Case Err

        Case 17840, 17861
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 17841 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = Vendedor.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 17842 'Não encontrou código do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Prenche objVendedor com codigo
                objVendedor.iCodigo = CDbl(Vendedor.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 17860
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158809)

    End Select

    Exit Function

End Function

Private Function Traz_Debito_Tela(objDebitoRecCli As ClassDebitoRecCli) As Long
'Traz os dados de debitos receber cliente para a Tela

Dim lErro As Long
Dim colInfoComissao As New colInfoComissao
Dim objInfoComissao As ClassInfoComissao
Dim iLinha As Integer
Dim iIndice As Integer, Cancel As Boolean

On Error GoTo Erro_Traz_Debito_Tela

    'Coloca os dados do debito a receber cliente na tela
    Cliente.Text = objDebitoRecCli.lCliente
    Call Cliente_Validate(Cancel)

    Filial.Text = objDebitoRecCli.iFilial
    Call Filial_Validate(bSGECancelDummy)

    Tipo.Text = objDebitoRecCli.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)
    
    NumTitulo.Text = objDebitoRecCli.lNumTitulo

    Call DateParaMasked(DataEmissao, objDebitoRecCli.dtDataEmissao)

    ValorTotal.Text = objDebitoRecCli.dValorTotal
    ValorICMS.Text = objDebitoRecCli.dValorICMS
    ValorICMSSubst.Text = objDebitoRecCli.dValorICMSSubst
    ValorProdutos.Text = objDebitoRecCli.dValorProdutos
    ValorIRRF.Text = objDebitoRecCli.dValorIRRF
    ValorFrete.Text = objDebitoRecCli.dValorFrete
    ValorSeguro.Text = objDebitoRecCli.dValorSeguro
    ValorIPI.Text = objDebitoRecCli.dValorIPI
    OutrasDespesas.Text = objDebitoRecCli.dOutrasDespesas
    PISRetido.Text = objDebitoRecCli.dPISRetido
    COFINSRetido.Text = objDebitoRecCli.dCOFINSRetido
    CSLLRetido.Text = objDebitoRecCli.dCSLLRetido
    Observacao.Text = objDebitoRecCli.sObservacao
    
    Saldo.Caption = Format(objDebitoRecCli.dSaldo, "STANDARD")
    glNumIntDoc = objDebitoRecCli.lNumIntDoc

    'Lê as comissões vinculadas ao debito
    lErro = CF("Comissoes_Le_DebRecCli", objDebitoRecCli, colInfoComissao)
    If lErro <> SUCESSO Then Error 17843

    Call Grid_Limpa(objGridComissoes)

    If colInfoComissao.Count > NUM_MAXIMO_COMISSOES Then Error 17844

    iLinha = 0

    'Preenche as linhas do Grid Comissoes com os dados de cada comissao
    For Each objInfoComissao In colInfoComissao

        iLinha = iLinha + 1

        GridComissoes.TextMatrix(iLinha, iGrid_Vendedor_Col) = objInfoComissao.sVendedorNomeRed
        GridComissoes.TextMatrix(iLinha, iGrid_Percentual_Col) = Format(objInfoComissao.dPercentual, "Percent")
        GridComissoes.TextMatrix(iLinha, iGrid_ValorBase_Col) = Format(objInfoComissao.dValorBase, "Standard")
        GridComissoes.TextMatrix(iLinha, iGrid_Valor_Col) = Format(Abs(objInfoComissao.dValor), "Standard")

    Next

    'Faz o número de linhas existentes do Grid ser igual ao número de comissoes
    objGridComissoes.iLinhasExistentes = iLinha

    'Faz refresh nas checkboxes
    Call Grid_Refresh_Checkbox(objGridComissoes)

    'Chama SomaPercentual
    lErro = Soma_Percentual()
    If lErro <> SUCESSO Then Error 17951

    'Chama SomaValor
    lErro = Soma_Valor()
    If lErro <> SUCESSO Then Error 17952
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objDebitoRecCli.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then Error 39680

    iAlterado = 0
    iValorAlterado = 0

    Traz_Debito_Tela = SUCESSO

    Exit Function

Erro_Traz_Debito_Tela:

    Traz_Debito_Tela = Err

    Select Case Err

        Case 17843, 17951, 17952

        Case 17844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_COMISSOES_ULTRAPASSADO", Err, colInfoComissao.Count, NUM_MAXIMO_COMISSOES)
        
        Case 39680
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158810)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objDebitoRecCli As ClassDebitoRecCli, colInfoComissao As colInfoComissao) As Long
'Move os dados da Tela para objDebitosRecCli e colInfoComissao

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_Tela_Memoria

    objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(Cliente.Text)) > 0 Then
        objCliente.sNomeReduzido = Cliente.Text

        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 17850

        If lErro <> SUCESSO Then Error 17851

        objDebitoRecCli.lCliente = objCliente.lCodigo
    End If

    If Len(Trim(Filial.Text)) > 0 Then
        objDebitoRecCli.iFilial = Codigo_Extrai(Filial.Text)
    End If

    If Len(Trim(Tipo.Text)) > 0 Then objDebitoRecCli.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)

    If Len(Trim(NumTitulo.ClipText)) > 0 Then objDebitoRecCli.lNumTitulo = CLng(NumTitulo.Text)

    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objDebitoRecCli.dtDataEmissao = DATA_NULA
    Else
        objDebitoRecCli.dtDataEmissao = CDate(DataEmissao.Text)
    End If

    If Len(Trim(ValorTotal.ClipText)) > 0 Then objDebitoRecCli.dValorTotal = CDbl(ValorTotal.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then objDebitoRecCli.dValorSeguro = CDbl(ValorSeguro)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then objDebitoRecCli.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(OutrasDespesas.ClipText)) > 0 Then objDebitoRecCli.dOutrasDespesas = CDbl(OutrasDespesas.Text)
    If Len(Trim(ValorProdutos.ClipText)) > 0 Then objDebitoRecCli.dValorProdutos = CDbl(ValorProdutos.Text)
    If Len(Trim(ValorICMS.ClipText)) > 0 Then objDebitoRecCli.dValorICMS = CDbl(ValorICMS.Text)
    If Len(Trim(ValorICMSSubst.ClipText)) > 0 Then objDebitoRecCli.dValorICMSSubst = CDbl(ValorICMSSubst.Text)
    If Len(Trim(ValorIPI.ClipText)) > 0 Then objDebitoRecCli.dValorIPI = CDbl(ValorIPI.Text)
    If Len(Trim(ValorIRRF.ClipText)) > 0 Then objDebitoRecCli.dValorIRRF = CDbl(ValorIRRF.Text)
    If Len(Trim(PISRetido.ClipText)) > 0 Then objDebitoRecCli.dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(COFINSRetido.ClipText)) > 0 Then objDebitoRecCli.dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.ClipText)) > 0 Then objDebitoRecCli.dCSLLRetido = CDbl(CSLLRetido.Text)
    
    objDebitoRecCli.sObservacao = Observacao.Text

    'Move para colInfoComissao os dados do Grid Comissoes
    lErro = Move_GridComissoes_Memoria(colInfoComissao)
    If lErro <> SUCESSO Then Error 17852

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 17850, 17852

        Case 17851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158811)

    End Select

    Exit Function

End Function

Private Function Move_GridComissoes_Memoria(colInfoComissao As colInfoComissao) As Long
'Move para a memória os dados existentes no Grid

Dim lErro As Long
Dim iIndice As Integer
Dim lComprimento As Long
Dim objInfoComissao As New ClassInfoComissao
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Move_GridComissoes_Memoria

    For iIndice = 1 To objGridComissoes.iLinhasExistentes

        Set objInfoComissao = New ClassInfoComissao

        If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col))) > 0 Then
            'Se Vendedor estiver preenchido, verifica se existe
            objVendedor.sNomeReduzido = (GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col))

            lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
            If lErro <> SUCESSO And lErro <> 25008 Then Error 17853

            If lErro = 25008 Then Error 17854

            objInfoComissao.iCodVendedor = objVendedor.iCodigo

        End If

        lComprimento = Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Col)))
        If lComprimento > 0 Then objInfoComissao.dPercentual = PercentParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Col))

        objInfoComissao.dValorBase = StrParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_ValorBase_Col))

        objInfoComissao.dValor = StrParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Col))

        With objInfoComissao
            colInfoComissao.Add .lNumIntCom, .iTipoTitulo, .lNumIntDoc, .iCodVendedor, .dtDataBaixa, .dPercentual, .dValorBase, .dValor, .iStatus, .iFilialEmpresa, .sVendedorNomeRed, .dtDataGeracao
        End With
    Next

    Move_GridComissoes_Memoria = SUCESSO

    Exit Function

Erro_Move_GridComissoes_Memoria:

    Move_GridComissoes_Memoria = Err

    Select Case Err

        Case 17853

        Case 17854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158812)

    End Select

    Exit Function

End Function

Private Function Soma_Percentual() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dSomaPercentual As Double
Dim lComprimento As Long

On Error GoTo Erro_Soma_Percentual

    dSomaPercentual = 0

    'Loop no GridComissao
    For iIndice = 1 To objGridComissoes.iLinhasExistentes
        lComprimento = Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Col)))
        'Verifica se Percentual da Comissão está preenchido
        If lComprimento > 0 Then
            'Acumula Percentual em dSomaPercentual
            dSomaPercentual = dSomaPercentual + CDbl(left(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Col), lComprimento - 1))
        End If
    Next

    'Mostra na tela o Total Percentual
    TotalPercentual.Caption = Format(dSomaPercentual, "#0.#0\%")

    Soma_Percentual = SUCESSO

    Exit Function

Erro_Soma_Percentual:

    Soma_Percentual = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158813)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorBase(objGridInt As AdmGrid) As Long
'Faz a crítica da celula ValorBase do grid que está deixando de ser a corrente

Dim lErro As Long
Dim cPercentual As Currency
Dim dValorBase As Double
Dim dValorComissao As Double
Dim dValorTotal As Double
Dim lComprimento As Long

On Error GoTo Erro_Saida_Celula_ValorBase

    Set objGridInt.objControle = ValorBase

    'Verifica se valor base está preenchido
    If Len(Trim(ValorBase.ClipText)) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_Positivo_Critica(ValorBase.Text)
        If lErro <> SUCESSO Then Error 17865

        dValorBase = CDbl(ValorBase.Text)
        
        If Len(Trim(ValorTotal.Text)) <> 0 Then dValorTotal = CDbl(ValorTotal.Text)
        
        If dValorBase > dValorTotal Then Error 19152

        'Mostra na tela o ValorBase formatado
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col) = Format(dValorBase, "Fixed")

        'Verifica se percentual comissão está preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Col))) > 0 Then
        
            lComprimento = Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Col)))
        
            If lComprimento > 0 Then cPercentual = CDbl(left(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Col), lComprimento - 1))

            'Calcula o valor da comissao
            dValorComissao = (cPercentual * dValorBase) / 100

            'Mostra na tela o valor da comissao
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Col) = Format(dValorComissao, "Standard")
            
        Else
            'Verifica se valor comissão está preenchido
            If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Col))) Then
            
                dValorComissao = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Col))
                
                cPercentual = (dValorComissao / dValorBase) * 100
                
                'Mostra o percentual da comissão na tela
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Col) = Format(cPercentual, "#0.#0\%")
                
            End If
            
        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 17866

    'Chama SomaValor
    lErro = Soma_Valor()
    If lErro <> SUCESSO Then Error 17950

    Saida_Celula_ValorBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBase:

    Saida_Celula_ValorBase = Err

    Select Case Err

        Case 17865, 17866
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 17950

        Case 19152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORES_TOTAL_BASE", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158814)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim cPercentual As Currency
Dim dValorBase As Double
Dim dValorComissao As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    'Verifica se valor está preenchido
    If Len(Trim(Valor.ClipText)) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 17867

        dValorComissao = CDbl(Valor.Text)

        'Mostra na tela o Valor
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Col) = Format(dValorComissao, "Fixed")

        'Verifica se valor base correspondente está preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))
            
            'Valor base não pode ser menor que o da comissão
            If dValorBase < dValorComissao Then Error 19323
            
            'Recalcula Percentual
            If dValorComissao <> 0 And dValorBase <> 0 Then
                
                cPercentual = (dValorComissao / dValorBase) * 100
                'Mostra o percentual da comissão na tela
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Col) = Format(cPercentual, "#0.#0\%")
                
            End If

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 17866

    'Chama SomaPercentual
    lErro = Soma_Percentual()
    If lErro <> SUCESSO Then Error 17949

    'Chama SomaValor
    lErro = Soma_Valor()
    If lErro <> SUCESSO Then Error 17868

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 17867, 17866
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 17868, 17949
        
        Case 19323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORES_COMISSAO_BASE", Err, dValorBase, dValorComissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158815)

    End Select

    Exit Function

End Function

Private Function Soma_Valor() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dSomaValor As Double

On Error GoTo Erro_Soma_Valor

    dSomaValor = 0

    'Loop no GridComissao
    For iIndice = 1 To objGridComissoes.iLinhasExistentes

        'Verifica se Valor da Comissão está preenchido
        If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Col))) > 0 Then

            'Acumula Valor em dSomaValor
            dSomaValor = dSomaValor + CDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Col))

        End If

    Next

    'Mostra na tela o Total Valor
    TotalValor.Caption = Format(dSomaValor, "Fixed")

    Soma_Valor = SUCESSO

    Exit Function

Erro_Soma_Valor:

    Soma_Valor = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158816)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "DebitosRecCli"

    'Lê os dados da Tela Debitos a Receber Cliente
    lErro = Move_Tela_Memoria(objDebitoRecCli, colInfoComissao)
    If lErro <> SUCESSO Then Error 17869

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", CLng(0), 0, "NumIntDoc"
    colCampoValor.Add "Cliente", objDebitoRecCli.lCliente, 0, "Cliente"
    colCampoValor.Add "Filial", objDebitoRecCli.iFilial, 0, "Filial"
    colCampoValor.Add "SiglaDocumento", objDebitoRecCli.sSiglaDocumento, STRING_DEB_REC_SIGLA, "SiglaDocumento"
    colCampoValor.Add "NumTitulo", objDebitoRecCli.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objDebitoRecCli.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorTotal", objDebitoRecCli.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorSeguro", objDebitoRecCli.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorFrete", objDebitoRecCli.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "OutrasDespesas", objDebitoRecCli.dOutrasDespesas, 0, "OutrasDespesas"
    colCampoValor.Add "ValorProdutos", objDebitoRecCli.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorICMS", objDebitoRecCli.dValorICMS, 0, "ValorICMS"
    colCampoValor.Add "ValorICMSSubst", objDebitoRecCli.dValorICMS, 0, "ValorICMSSubst"
    colCampoValor.Add "ValorIPI", objDebitoRecCli.dValorIPI, 0, "ValorIPI"
    colCampoValor.Add "ValorIRRF", objDebitoRecCli.dValorIRRF, 0, "ValorIRRF"
    colCampoValor.Add "PISRetido", objDebitoRecCli.dPISRetido, 0, "PISRetido"
    colCampoValor.Add "COFINSRetido", objDebitoRecCli.dCOFINSRetido, 0, "COFINSRetido"
    colCampoValor.Add "CSLLRetido", objDebitoRecCli.dCSLLRetido, 0, "CSLLRetido"

    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 17869

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158817)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objDebitoRecCli As New ClassDebitoRecCli

On Error GoTo Erro_Tela_Preenche

    objDebitoRecCli.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objDebitoRecCli.lNumIntDoc <> 0 Then

        'Carrega objDebitoRecCli com os dados passados em colCampoValor
        objDebitoRecCli.lCliente = colCampoValor.Item("Cliente").vValor
        objDebitoRecCli.iFilial = colCampoValor.Item("Filial").vValor
        objDebitoRecCli.sSiglaDocumento = colCampoValor.Item("SiglaDocumento").vValor
        objDebitoRecCli.lNumTitulo = colCampoValor.Item("NumTitulo").vValor
        objDebitoRecCli.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objDebitoRecCli.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objDebitoRecCli.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
        objDebitoRecCli.dValorFrete = colCampoValor.Item("ValorFrete").vValor
        objDebitoRecCli.dOutrasDespesas = colCampoValor.Item("OutrasDespesas").vValor
        objDebitoRecCli.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
        objDebitoRecCli.dValorICMS = colCampoValor.Item("ValorICMS").vValor
        objDebitoRecCli.dValorICMSSubst = colCampoValor.Item("ValorICMSSubst").vValor
        objDebitoRecCli.dValorIPI = colCampoValor.Item("ValorIPI").vValor
        objDebitoRecCli.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        objDebitoRecCli.dPISRetido = colCampoValor.Item("PISRetido").vValor
        objDebitoRecCli.dCOFINSRetido = colCampoValor.Item("COFINSRetido").vValor
        objDebitoRecCli.dCSLLRetido = colCampoValor.Item("CSLLRetido").vValor

        'Verifica se o debito a receber existe
        lErro = CF("DebitoReceber_Le", objDebitoRecCli)
        If lErro <> SUCESSO And lErro <> 17835 Then Error 17870

        lErro = Traz_Debito_Tela(objDebitoRecCli)
        If lErro <> SUCESSO Then Error 17870

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 17870

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158818)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValorProdutos As Double
Dim dValorICMSSubst As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorIPI As Double
Dim dOutrasDespesas As Double
Dim dValorIRRF As Double
Dim dSoma As Double
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos obrigatórios estão preenchidos
    If Len(Trim(Cliente.Text)) = 0 Then Error 17872
    If Len(Trim(Filial.Text)) = 0 Then Error 17873
    If Len(Trim(Tipo.Text)) = 0 Then Error 17874
    If Len(Trim(NumTitulo.Text)) = 0 Then Error 17875
    If Len(Trim(ValorTotal.ClipText)) = 0 Then Error 17876
    If Len(Trim(ValorProdutos.ClipText)) = 0 Then Error 17877

    'Conversão dos valores para Double
    If Len(Trim(ValorProdutos.Text)) > 0 Then dValorProdutos = CDbl(ValorProdutos.Text)
    If Len(Trim(ValorICMSSubst.Text)) > 0 Then dValorICMSSubst = CDbl(ValorICMSSubst)
    If Len(Trim(ValorFrete.Text)) > 0 Then dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.Text)) > 0 Then dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(OutrasDespesas.Text)) > 0 Then dOutrasDespesas = CDbl(OutrasDespesas.Text)
    If Len(Trim(ValorIPI.Text)) > 0 Then dValorIPI = CDbl(ValorIPI.Text)

    'Soma dos valores
    dSoma = dValorProdutos + dValorICMSSubst + dValorFrete + dValorSeguro + dOutrasDespesas + dValorIPI

    'Verifica se Soma dos Valores é igual ao ValorTotal
    If Abs(dSoma - CDbl(ValorTotal.Text)) > DELTA_VALORMONETARIO Then Error 17878

    'Move dados da Tela para objDebitoRecCli e colComissao
    lErro = Move_Tela_Memoria(objDebitoRecCli, colInfoComissao)
    If lErro <> SUCESSO Then Error 17879

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    If objDebitoRecCli.dtDataEmissao <> DATA_NULA Then
        lErro = objContabil.Contabil_Testa_Data(objDebitoRecCli.dtDataEmissao)
        If lErro <> SUCESSO Then Error 20830
    End If

    'Grava Débito a Receber Cliente no BD (inclusive os daos contábeis) (contabilidade)
    lErro = CF("DebitoRecCli_Grava", objDebitoRecCli, colInfoComissao, objContabil)
    If lErro <> SUCESSO Then Error 17880

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function


Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 17873
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 17874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", Err)

        Case 17875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case 17876
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NAO_INFORMADO", Err)

        Case 17877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPRODUTOS_NAO_INFORMADO", Err)

        Case 17878
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_INVALIDO", Err, ValorTotal.Text, dSoma)

        Case 17879, 17880, 20830

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158819)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_DebitoRecCli()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Chama função que limpa TextBoxes e MaskedEdits da Tela
    Call Limpa_Tela(Me)

    'Limpa os campos não são limpos pela função acima
    Filial.Clear
    Tipo.Text = ""
    Saldo.Caption = ""
    glNumIntDoc = 0

    Call Grid_Limpa(objGridComissoes)

    TotalPercentual.Caption = ""
    TotalValor.Caption = ""
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0

End Sub

Private Sub Vendedores_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    If GridComissoes.Row >= GridComissoes.FixedRows Then

        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Vendedor_Col))) > 0 Then

            'Preenche NomeReduzido com o vendedor da tela
            objVendedor.sNomeReduzido = GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Vendedor_Col)
        End If
    End If

    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

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

Dim lErro As Long
Dim objCliente As New ClassCliente, objTipoCliente As New ClassTipoCliente
Dim objFilial As New ClassFilialCliente
Dim sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case CLIENTE_COD
            
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 56510
                
                objMnemonicoValor.colValor.Add objCliente.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case CLIENTE_NOME
        
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 56511
            
                objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case FILIAL_COD
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                objMnemonicoValor.colValor.Add objFilial.iCodFilial
            
            Else
                
                objMnemonicoValor.colValor.Add 0
            
            End If
            
        Case FILIAL_NOME_RED
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 56512
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 56513
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56514
                
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
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 56515
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case TIPO1
            If Len(Tipo.Text) > 0 Then
                objMnemonicoValor.colValor.Add Tipo.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case NUM_TITULO
            If Len(NumTitulo.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NumTitulo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case DATA_EMISSAO
            If Len(DataEmissao.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case VALOR_TOTAL
            If Len(Trim(ValorTotal.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorTotal.Text)
            Else
                objMnemonicoValor.colValor.Add 0
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
            
        Case VALOR_PRODUTOS
            If Len(Trim(ValorProdutos.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorProdutos.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_IRRF
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
            
        Case OUTRAS_DESPESAS
            If Len(Trim(OutrasDespesas.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(OutrasDespesas.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CONTA_TIPO_CLIENTE
            
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 56516
                
                objTipoCliente.iCodigo = objCliente.iTipo
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO Then Error 56517
                
                If objTipoCliente.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTipoCliente.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56518
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case Else
            Error 39681
            
    End Select
    
    Calcula_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico:
    
    Calcula_Mnemonico = Err
    
    Select Case Err
    
        Case 56510 To 56518
        
        Case 39681
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158820)

    End Select
    
    Exit Function
    
End Function

Private Function Vendedor_Linha_Preenche(objVendedor As ClassVendedor) As Long
'gera defaults p/linha de grid de comissoes apos preenchimento do vendedor

Dim lErro As Long, iCodFilial As Integer, objComissao As New ClassComissao
Dim iIndice As Integer

On Error GoTo Erro_Vendedor_Linha_Preenche

    If UCase(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Vendedor_Col)) <> UCase(objVendedor.sNomeReduzido) Then
    
        'Loop no GridComissoes
        For iIndice = 1 To objGridComissoes.iLinhasExistentes
    
            'Verifica se Vendedor comparece em outra linha
            If iIndice <> GridComissoes.Row Then If UCase(GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)) = UCase(objVendedor.sNomeReduzido) Then Error 59210
    
        Next
    
        iCodFilial = Codigo_Extrai(Filial.Text)
        lErro = ComissaoAutomatica_Obter_Debito(Cliente.Text, iCodFilial, objVendedor, objComissao, StrParaDbl(ValorTotal.Text), StrParaDbl(ValorFrete.Text), StrParaDbl(ValorSeguro.Text), StrParaDbl(OutrasDespesas.Text), StrParaDbl(ValorIPI.Text), StrParaDbl(ValorICMSSubst.Text))
        If lErro <> SUCESSO Then Error 59211
        
        If objComissao.iCodVendedor <> 0 Then
            
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
            
            Vendedor.PromptInclude = False
            Vendedor.Text = objVendedor.sNomeReduzido
            Vendedor.PromptInclude = True
            
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Col) = Format(objComissao.dPercentual, "Percent")
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Col) = Format(objComissao.dValor, "Standard")
            
            If GridComissoes.Row > objGridComissoes.iLinhasExistentes Then objGridComissoes.iLinhasExistentes = objGridComissoes.iLinhasExistentes + 1
        
            Call Soma_Percentual
            Call Soma_Valor
    
        End If
        
    End If
    
    Vendedor_Linha_Preenche = SUCESSO
    
    Exit Function
     
Erro_Vendedor_Linha_Preenche:

    Vendedor_Linha_Preenche = Err
    
    Select Case Err
          
        Case 59210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", Err, objVendedor.sNomeReduzido)
    
        Case 59211
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158821)
     
    End Select
     
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_DEVOL_DEB_CLIENTES_ID
    Set Form_Load_Ocx = Me
    Caption = "Devoluções / Débitos de Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "DebitosReceb"
    
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
        
        If Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call TipoLabel_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call Vendedores_Click
        End If
    
    End If
    
    
End Sub

'?????? transferir p/crfatcritica
Private Function ComissaoAutomatica_Obter_Debito(sClienteNomeRed As String, iCodFilial As Integer, objVendedor As ClassVendedor, objComissao As ClassComissao, dValorTotal As Double, dValorFrete As Double, dValorSeguro As Double, dValorOutras As Double, dValorIPI As Double, Optional ByVal dValorICMSST As Double = 0) As Long
'retorna em objVendedor e objComissao valores default para geracao de comissao automatica à partir do cliente e filial informados
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim lErro As Long, objComissaoNF As New ClassComissaoNF

On Error GoTo Erro_ComissaoAutomatica_Obter_Debito

    objComissao.dValorBase = 0
    objComissao.dValor = 0
    objComissao.dPercentual = 0
    objComissao.iCodVendedor = 0
    
    lErro = CF("ComissaoAutomatica_Obter_Info", sClienteNomeRed, iCodFilial, objVendedor, objComissaoNF)
    If lErro <> SUCESSO Then Error 59208
    
    'Verifica se achou o Vendedor
    If objVendedor.iCodigo <> 0 Then
    
        objComissao.iCodVendedor = objVendedor.iCodigo
    
        If objVendedor.iComissaoSobreTotal = 0 Then
            objComissao.dValorBase = Round(dValorTotal - IIf(objVendedor.iComissaoFrete = 0, dValorFrete, 0) - IIf(objVendedor.iComissaoSeguro = 0, dValorSeguro, 0) - IIf(objVendedor.iComissaoIPI = 0, dValorIPI, 0) - IIf(objVendedor.iComissaoICM = 0, dValorOutras, 0) - dValorICMSST, 2)
        Else
            objComissao.dValorBase = dValorTotal
        End If
        
        objComissao.dValor = Round(objComissao.dValorBase * objComissaoNF.dPercentual, 2)
        objComissao.dPercentual = objComissaoNF.dPercentual
        
    End If
    
    ComissaoAutomatica_Obter_Debito = SUCESSO
     
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Debito:

    ComissaoAutomatica_Obter_Debito = Err
     
    Select Case Err
          
        Case 59208
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158822)
     
    End Select
     
    Exit Function

End Function



Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
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

Private Sub TotalPercentual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentual, Source, X, Y)
End Sub

Private Sub TotalPercentual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentual, Button, Shift, X, Y)
End Sub

Private Sub TotalValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValor, Source, X, Y)
End Sub

Private Sub TotalValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValor, Button, Shift, X, Y)
End Sub

Private Sub TotalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalLabel, Source, X, Y)
End Sub

Private Sub TotalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalLabel, Button, Shift, X, Y)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158823)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158824)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158825)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134025

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134025

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158826)

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
    colSelecao.Add MOTIVO_DEBITO_CLIENTE
    colSelecao.Add glNumIntDoc

    'Abre o Browse de Antecipações de recebimento de uma Filial
    Call Chama_Tela("BaixasRecLista", colSelecao, Nothing, Nothing, "NumIntBaixa IN (SELECT NumIntBaixa FROM BaixasRec WHERE Motivo = ? AND NumIntDoc = ? AND Status <> 5)")

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

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaPagOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5160
      Index           =   2
      Left            =   60
      TabIndex        =   48
      Top             =   705
      Visible         =   0   'False
      Width           =   9405
      Begin VB.Frame FramePagamento 
         Caption         =   "Cheques de Terceiros"
         Height          =   1665
         Index           =   3
         Left            =   0
         TabIndex        =   133
         Top             =   3495
         Visible         =   0   'False
         Width           =   9405
         Begin MSMask.MaskEdBox ValorJurosCT 
            Height          =   225
            Left            =   4785
            TabIndex        =   150
            Top             =   705
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox ValorBaixadoCT 
            Height          =   225
            Left            =   3975
            TabIndex        =   149
            Top             =   1080
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox ValorDescontoCT 
            Height          =   225
            Left            =   6855
            TabIndex        =   148
            Top             =   915
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox ValorMultaCT 
            Height          =   225
            Left            =   5610
            TabIndex        =   147
            Top             =   720
            Width           =   1080
            _ExtentX        =   1905
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
         Begin VB.TextBox ClienteCT 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5205
            TabIndex        =   144
            Top             =   1020
            Width           =   1500
         End
         Begin VB.TextBox FilialCT 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3015
            TabIndex        =   143
            Top             =   420
            Width           =   1500
         End
         Begin VB.TextBox BancoCT 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1815
            TabIndex        =   142
            Top             =   435
            Width           =   750
         End
         Begin VB.TextBox AgenciaCT 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   375
            TabIndex        =   141
            Top             =   450
            Width           =   795
         End
         Begin VB.TextBox ContaCorrenteCT 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3570
            TabIndex        =   140
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox NumeroCT 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2415
            TabIndex        =   139
            Top             =   945
            Width           =   990
         End
         Begin VB.TextBox DataDepositoCT 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   345
            TabIndex        =   138
            Top             =   690
            Width           =   1500
         End
         Begin MSMask.MaskEdBox ValorCT 
            Height          =   225
            Left            =   4665
            TabIndex        =   136
            Top             =   450
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox FilialEmpresaCT 
            Height          =   225
            Left            =   6615
            TabIndex        =   137
            Top             =   405
            Width           =   1125
            _ExtentX        =   1984
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
         Begin VB.CheckBox SelecionadoCT 
            Height          =   225
            Left            =   645
            TabIndex        =   134
            Top             =   1020
            Width           =   510
         End
         Begin MSFlexGridLib.MSFlexGrid GridChequePre 
            Height          =   1215
            Left            =   15
            TabIndex        =   135
            Top             =   195
            Width           =   9360
            _ExtentX        =   16510
            _ExtentY        =   2143
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   6615
            TabIndex        =   146
            Top             =   1500
            Width           =   510
         End
         Begin VB.Label TotalCheque 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   7185
            TabIndex        =   145
            Top             =   1455
            Width           =   1560
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Cr�ditos"
         Height          =   1665
         Index           =   2
         Left            =   0
         TabIndex        =   59
         Top             =   3495
         Visible         =   0   'False
         Width           =   9405
         Begin VB.CheckBox Selecionado 
            Height          =   220
            Left            =   6540
            TabIndex        =   60
            Top             =   165
            Width           =   585
         End
         Begin MSMask.MaskEdBox SaldoCredito 
            Height          =   225
            Left            =   4305
            TabIndex        =   61
            Top             =   180
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox ValorCredito 
            Height          =   225
            Left            =   3120
            TabIndex        =   62
            Top             =   165
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox NumTitulo 
            Height          =   225
            Left            =   2355
            TabIndex        =   63
            Top             =   150
            Width           =   645
            _ExtentX        =   1138
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
         Begin MSMask.MaskEdBox SiglaDocumento 
            Height          =   225
            Left            =   1830
            TabIndex        =   64
            Top             =   165
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   225
            Left            =   735
            TabIndex        =   65
            Top             =   165
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialEmpresaCR 
            Height          =   225
            Left            =   6885
            TabIndex        =   66
            Top             =   165
            Width           =   1170
            _ExtentX        =   2064
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
         Begin MSFlexGridLib.MSFlexGrid GridCreditos 
            Height          =   1245
            Left            =   165
            TabIndex        =   67
            Top             =   270
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2196
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Dados do Pagamento"
         Height          =   1665
         Index           =   0
         Left            =   0
         TabIndex        =   50
         Top             =   3495
         Width           =   9405
         Begin VB.ComboBox ContaCorrente 
            Height          =   315
            Left            =   1080
            TabIndex        =   52
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox Portador 
            Height          =   315
            Left            =   4545
            Sorted          =   -1  'True
            TabIndex        =   51
            Top             =   780
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Historico 
            Height          =   300
            Left            =   1080
            TabIndex        =   53
            Top             =   1185
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label ValorPago 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1080
            TabIndex        =   58
            Top             =   780
            Width           =   1680
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   450
            TabIndex        =   57
            Top             =   405
            Width           =   570
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Hist�rico:"
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
            Left            =   195
            TabIndex        =   56
            Top             =   1230
            Width           =   825
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Local de Pagto:"
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
            Left            =   3120
            TabIndex        =   55
            Top             =   810
            Width           =   1365
         End
         Begin VB.Label Label25 
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
            Left            =   510
            TabIndex        =   54
            Top             =   810
            Width           =   510
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Adiantamentos a Fornecedor"
         Height          =   1635
         Index           =   1
         Left            =   0
         TabIndex        =   68
         Top             =   3525
         Visible         =   0   'False
         Width           =   9405
         Begin VB.CheckBox SelecionadoPA 
            Height          =   220
            Left            =   330
            TabIndex        =   69
            Top             =   210
            Width           =   525
         End
         Begin MSMask.MaskEdBox CCIntNomeReduzido 
            Height          =   225
            Left            =   2025
            TabIndex        =   70
            Top             =   195
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MeioPagtoDescricao 
            Height          =   225
            Left            =   3030
            TabIndex        =   71
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroMP 
            Height          =   225
            Left            =   4155
            TabIndex        =   72
            Top             =   165
            Width           =   690
            _ExtentX        =   1217
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
         Begin MSMask.MaskEdBox DataMovimento 
            Height          =   225
            Left            =   990
            TabIndex        =   73
            Top             =   165
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorPA 
            Height          =   225
            Left            =   4860
            TabIndex        =   74
            Top             =   165
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox SaldoPA 
            Height          =   225
            Left            =   5880
            TabIndex        =   75
            Top             =   180
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox FilialEmpresaPA 
            Height          =   225
            Left            =   7635
            TabIndex        =   76
            Top             =   405
            Width           =   1245
            _ExtentX        =   2196
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
         Begin MSFlexGridLib.MSFlexGrid GridPagtosAntecipados 
            Height          =   1215
            Left            =   150
            TabIndex        =   77
            Top             =   240
            Width           =   8625
            _ExtentX        =   15214
            _ExtentY        =   2143
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.ComboBox Ordenacao 
         Height          =   315
         ItemData        =   "BaixaPagOcx.ctx":0000
         Left            =   6390
         List            =   "BaixaPagOcx.ctx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   105
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Baixa"
         Height          =   510
         Left            =   0
         TabIndex        =   93
         Top             =   2970
         Width           =   9405
         Begin VB.OptionButton Pagamento 
            Caption         =   "Cheques de Terc."
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
            Index           =   3
            Left            =   7155
            TabIndex        =   132
            Top             =   225
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Cr�dito / Devolu��o"
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
            Left            =   4715
            TabIndex        =   96
            Top             =   210
            Width           =   2055
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Adiantamento"
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
            Index           =   1
            Left            =   2800
            TabIndex        =   95
            Top             =   210
            Width           =   1530
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Pagamento em dinheiro"
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
            Index           =   0
            Left            =   45
            TabIndex        =   94
            Top             =   195
            Value           =   -1  'True
            Width           =   2370
         End
      End
      Begin VB.Frame FrameParcelas 
         Caption         =   "Parcelas em Aberto"
         Height          =   2565
         Left            =   0
         TabIndex        =   78
         Top             =   405
         Width           =   9405
         Begin VB.CheckBox Selecionada 
            Height          =   220
            Left            =   8025
            TabIndex        =   81
            Top             =   270
            Width           =   570
         End
         Begin MSMask.MaskEdBox DataEmissaoTitulo 
            Height          =   225
            Left            =   4530
            TabIndex        =   79
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   4485
            TabIndex        =   80
            Top             =   570
            Width           =   1290
            _ExtentX        =   2275
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
         Begin MSMask.MaskEdBox ValorAPagar 
            Height          =   225
            Left            =   5595
            TabIndex        =   82
            Top             =   360
            Width           =   1290
            _ExtentX        =   2275
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
         Begin MSMask.MaskEdBox ValorJuros 
            Height          =   225
            Left            =   7140
            TabIndex        =   83
            Top             =   255
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox ValorBaixado 
            Height          =   225
            Left            =   3930
            TabIndex        =   84
            Top             =   255
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSMask.MaskEdBox ValorMulta 
            Height          =   225
            Left            =   6255
            TabIndex        =   85
            Top             =   465
            Width           =   975
            _ExtentX        =   1720
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
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   225
            Left            =   5070
            TabIndex        =   86
            Top             =   255
            Width           =   1140
            _ExtentX        =   2011
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
            Left            =   360
            TabIndex        =   87
            Top             =   225
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Saldo 
            Height          =   225
            Left            =   3000
            TabIndex        =   88
            Top             =   255
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   1785
            TabIndex        =   89
            Top             =   240
            Width           =   795
            _ExtentX        =   1402
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
            Left            =   1275
            TabIndex        =   90
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
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
            Left            =   2640
            TabIndex        =   91
            Top             =   255
            Width           =   630
            _ExtentX        =   1111
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
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1350
            Left            =   15
            TabIndex        =   92
            Top             =   225
            Width           =   9360
            _ExtentX        =   16510
            _ExtentY        =   2381
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox FilialFornItem 
            Height          =   225
            Left            =   1530
            TabIndex        =   152
            Top             =   135
            Width           =   2025
            _ExtentX        =   3572
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornItem 
            Height          =   225
            Left            =   0
            TabIndex        =   153
            Top             =   0
            Width           =   2025
            _ExtentX        =   3572
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
            PromptChar      =   " "
         End
      End
      Begin VB.CommandButton BotaoConsultaDocOriginal 
         Height          =   450
         Left            =   -10000
         Picture         =   "BaixaPagOcx.ctx":0087
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Consulta o documento original de uma parcela, adiantamento ou cr�dito / devolu��o."
         Top             =   30
         Visible         =   0   'False
         Width           =   1065
      End
      Begin MSMask.MaskEdBox NomePortador 
         Height          =   225
         Left            =   6510
         TabIndex        =   97
         Top             =   615
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cobranca 
         Height          =   225
         Left            =   4020
         TabIndex        =   98
         Top             =   1425
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataBaixa 
         Height          =   300
         Left            =   2610
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   90
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBaixa 
         Height          =   300
         Left            =   1515
         TabIndex        =   99
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FilialEmpresa 
         Height          =   255
         Left            =   7560
         TabIndex        =   102
         Top             =   600
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordena��o:"
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
         Left            =   5385
         TabIndex        =   154
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label TotalBaixar 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3690
         TabIndex        =   105
         Top             =   120
         Width           =   1560
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
         Height          =   195
         Left            =   3120
         TabIndex        =   104
         Top             =   150
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data da Baixa:"
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
         Left            =   165
         TabIndex        =   103
         Top             =   150
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   1
      Left            =   150
      TabIndex        =   106
      Top             =   720
      Width           =   9120
      Begin VB.Frame Frame9 
         Caption         =   "Filtros"
         Height          =   3585
         Left            =   255
         TabIndex        =   112
         Top             =   1455
         Width           =   8355
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de Documento"
            Height          =   1410
            Left            =   390
            TabIndex        =   155
            Top             =   1950
            Width           =   4935
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
               Left            =   90
               TabIndex        =   158
               Top             =   960
               Width           =   1050
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
               Left            =   75
               TabIndex        =   157
               Top             =   360
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.ComboBox TipoDocSeleciona 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "BaixaPagOcx.ctx":0F91
               Left            =   1140
               List            =   "BaixaPagOcx.ctx":0F93
               Style           =   2  'Dropdown List
               TabIndex        =   156
               Top             =   930
               Width           =   3510
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "N� do T�tulo"
            Height          =   1575
            Left            =   5790
            TabIndex        =   127
            Top             =   270
            Width           =   2175
            Begin MSMask.MaskEdBox TituloInic 
               Height          =   300
               Left            =   720
               TabIndex        =   128
               Top             =   435
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TituloFim 
               Height          =   300
               Left            =   735
               TabIndex        =   129
               Top             =   960
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
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
               Left            =   315
               TabIndex        =   131
               Top             =   1005
               Width           =   375
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
               Left            =   360
               TabIndex        =   130
               Top             =   480
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data de Vencimento"
            Height          =   1575
            Left            =   3150
            TabIndex        =   120
            Top             =   270
            Width           =   2175
            Begin MSComCtl2.UpDown UpDownVencInic 
               Height          =   300
               Left            =   1695
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   480
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencInic 
               Height          =   300
               Left            =   630
               TabIndex        =   122
               Top             =   480
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
               Left            =   1695
               TabIndex        =   123
               TabStop         =   0   'False
               Top             =   990
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencFim 
               Height          =   300
               Left            =   615
               TabIndex        =   124
               Top             =   990
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label20 
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
               Left            =   210
               TabIndex        =   126
               Top             =   1020
               Width           =   375
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
               Left            =   240
               TabIndex        =   125
               Top             =   510
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data de Emiss�o"
            Height          =   1575
            Left            =   390
            TabIndex        =   113
            Top             =   270
            Width           =   2175
            Begin MSComCtl2.UpDown UpDownEmissaoInic 
               Height          =   300
               Left            =   1725
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   450
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoInic 
               Height          =   300
               Left            =   660
               TabIndex        =   115
               Top             =   465
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoFim 
               Height          =   300
               Left            =   1725
               TabIndex        =   116
               TabStop         =   0   'False
               Top             =   960
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoFim 
               Height          =   300
               Left            =   645
               TabIndex        =   117
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label16 
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
               Height          =   195
               Left            =   195
               TabIndex        =   119
               Top             =   1013
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
               Height          =   195
               Left            =   240
               TabIndex        =   118
               Top             =   495
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Fornecedor"
         Height          =   1005
         Left            =   255
         TabIndex        =   107
         Top             =   375
         Width           =   8355
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5475
            TabIndex        =   108
            Top             =   390
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1560
            TabIndex        =   109
            Top             =   397
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label FornecLabel 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   111
            Top             =   450
            Width           =   1035
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Filial:"
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
            Left            =   4920
            TabIndex        =   110
            Top             =   450
            Width           =   465
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   3
      Left            =   150
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   9120
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4680
         TabIndex        =   151
         Tag             =   "1"
         Top             =   1440
         Width           =   870
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
         TabIndex        =   18
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descri��o do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   13
         Top             =   3450
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   17
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   16
            Top             =   285
            Width           =   3720
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
            TabIndex        =   15
            Top             =   300
            Width           =   570
         End
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
            TabIndex        =   14
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   12
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   11
         Top             =   2175
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   10
         Top             =   2565
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
         Left            =   7740
         TabIndex        =   9
         Top             =   0
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padr�o"
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
         TabIndex        =   6
         Top             =   345
         Width           =   2700
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4920
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   47
         Top             =   165
         Width           =   450
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
         TabIndex        =   46
         Top             =   165
         Width           =   1035
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
         TabIndex        =   45
         Top             =   555
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   44
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   43
         Top             =   3030
         Width           =   1155
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
         TabIndex        =   42
         Top             =   3045
         Width           =   615
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
         TabIndex        =   41
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
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
         TabIndex        =   40
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Hist�ricos"
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
         TabIndex        =   39
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lan�amentos"
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
         TabIndex        =   38
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exerc�cio:"
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
         TabIndex        =   37
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   36
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   35
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Per�odo:"
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
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   33
         Top             =   120
         Width           =   1530
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
         TabIndex        =   32
         Top             =   165
         Width           =   720
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
         Left            =   6300
         TabIndex        =   31
         Top             =   690
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7815
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "BaixaPagOcx.ctx":0F95
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "BaixaPagOcx.ctx":1113
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "BaixaPagOcx.ctx":1645
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5610
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   9895
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "T�tulos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parcelas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabiliza��o"
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
Attribute VB_Name = "BaixaPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTBaixaPag
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoConsultaDocOriginal_Click()
    Call objCT.BotaoConsultaDocOriginal_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub CCIntNomeReduzido_GotFocus()
     Call objCT.CCIntNomeReduzido_GotFocus
End Sub

Private Sub CCIntNomeReduzido_KeyPress(KeyAscii As Integer)
     Call objCT.CCIntNomeReduzido_KeyPress(KeyAscii)
End Sub

Private Sub CCIntNomeReduzido_Validate(Cancel As Boolean)
     Call objCT.CCIntNomeReduzido_Validate(Cancel)
End Sub

Private Sub Cobranca_GotFocus()
     Call objCT.Cobranca_GotFocus
End Sub

Private Sub Cobranca_KeyPress(KeyAscii As Integer)
     Call objCT.Cobranca_KeyPress(KeyAscii)
End Sub

Private Sub Cobranca_Validate(Cancel As Boolean)
     Call objCT.Cobranca_Validate(Cancel)
End Sub

Private Sub ContaCorrente_Click()
     Call objCT.ContaCorrente_Click
End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)
     Call objCT.ContaCorrente_Validate(Cancel)
End Sub

Private Sub DataBaixa_Change()
     Call objCT.DataBaixa_Change
End Sub

Private Sub DataBaixa_GotFocus()
     Call objCT.DataBaixa_GotFocus
End Sub

Private Sub DataBaixa_Validate(Cancel As Boolean)
     Call objCT.DataBaixa_Validate(Cancel)
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.DataEmissao_KeyPress(KeyAscii)
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub DataMovimento_GotFocus()
     Call objCT.DataMovimento_GotFocus
End Sub

Private Sub DataMovimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataMovimento_KeyPress(KeyAscii)
End Sub

Private Sub DataMovimento_Validate(Cancel As Boolean)
     Call objCT.DataMovimento_Validate(Cancel)
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

Private Sub EmissaoFim_Change()
     Call objCT.EmissaoFim_Change
End Sub

Private Sub EmissaoFim_GotFocus()
     Call objCT.EmissaoFim_GotFocus
End Sub

Private Sub EmissaoFim_Validate(Cancel As Boolean)
     Call objCT.EmissaoFim_Validate(Cancel)
End Sub

Private Sub EmissaoInic_Change()
     Call objCT.EmissaoInic_Change
End Sub

Private Sub EmissaoInic_GotFocus()
     Call objCT.EmissaoInic_GotFocus
End Sub

Private Sub EmissaoInic_Validate(Cancel As Boolean)
     Call objCT.EmissaoInic_Validate(Cancel)
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub FilialEmpresa_GotFocus()
     Call objCT.FilialEmpresa_GotFocus
End Sub

Private Sub FilialEmpresa_KeyPress(KeyAscii As Integer)
     Call objCT.FilialEmpresa_KeyPress(KeyAscii)
End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)
     Call objCT.FilialEmpresa_Validate(Cancel)
End Sub

Private Sub FilialEmpresaCR_GotFocus()
     Call objCT.FilialEmpresaCR_GotFocus
End Sub

Private Sub FilialEmpresaCR_KeyPress(KeyAscii As Integer)
     Call objCT.FilialEmpresaCR_KeyPress(KeyAscii)
End Sub

Private Sub FilialEmpresaCR_Validate(Cancel As Boolean)
     Call objCT.FilialEmpresaCR_Validate(Cancel)
End Sub

Private Sub FilialEmpresaPA_GotFocus()
     Call objCT.FilialEmpresaPA_GotFocus
End Sub

Private Sub FilialEmpresaPA_KeyPress(KeyAscii As Integer)
     Call objCT.FilialEmpresaPA_KeyPress(KeyAscii)
End Sub

Private Sub FilialEmpresaPA_Validate(Cancel As Boolean)
     Call objCT.FilialEmpresaPA_Validate(Cancel)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Fornecedor_Change()
     Call objCT.Fornecedor_Change
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
     Call objCT.Fornecedor_Validate(Cancel)
End Sub

Private Sub FornecLabel_Click()
     Call objCT.FornecLabel_Click
End Sub

Private Sub GridCreditos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call objCT.GridCreditos_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub GridPagtosAntecipados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call objCT.GridPagtosAntecipados_MouseDown(Button, Shift, X, Y)
End Sub

Public Sub mnuGridConsultaDocOriginal_Click()
    Call objCT.mnuGridConsultaDocOriginal_Click
End Sub

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call objCT.GridParcelas_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub GridPagtosAntecipados_Click()
     Call objCT.GridPagtosAntecipados_Click
End Sub

Private Sub GridPagtosAntecipados_GotFocus()
     Call objCT.GridPagtosAntecipados_GotFocus
End Sub

Private Sub GridPagtosAntecipados_EnterCell()
     Call objCT.GridPagtosAntecipados_EnterCell
End Sub

Private Sub GridPagtosAntecipados_LeaveCell()
     Call objCT.GridPagtosAntecipados_LeaveCell
End Sub

Private Sub GridPagtosAntecipados_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridPagtosAntecipados_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridPagtosAntecipados_KeyPress(KeyAscii As Integer)
     Call objCT.GridPagtosAntecipados_KeyPress(KeyAscii)
End Sub

Private Sub GridPagtosAntecipados_Validate(Cancel As Boolean)
     Call objCT.GridPagtosAntecipados_Validate(Cancel)
End Sub

Private Sub GridPagtosAntecipados_RowColChange()
     Call objCT.GridPagtosAntecipados_RowColChange
End Sub

Private Sub GridPagtosAntecipados_Scroll()
     Call objCT.GridPagtosAntecipados_Scroll
End Sub

Private Sub GridCreditos_Click()
     Call objCT.GridCreditos_Click
End Sub

Private Sub GridCreditos_GotFocus()
     Call objCT.GridCreditos_GotFocus
End Sub

Private Sub GridCreditos_EnterCell()
     Call objCT.GridCreditos_EnterCell
End Sub

Private Sub GridCreditos_LeaveCell()
     Call objCT.GridCreditos_LeaveCell
End Sub

Private Sub GridCreditos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridCreditos_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridCreditos_KeyPress(KeyAscii As Integer)
     Call objCT.GridCreditos_KeyPress(KeyAscii)
End Sub

Private Sub GridCreditos_Validate(Cancel As Boolean)
     Call objCT.GridCreditos_Validate(Cancel)
End Sub

Private Sub GridCreditos_RowColChange()
     Call objCT.GridCreditos_RowColChange
End Sub

Private Sub GridCreditos_Scroll()
     Call objCT.GridCreditos_Scroll
End Sub

Private Sub Historico_Change()
     Call objCT.Historico_Change
End Sub

Private Sub MeioPagtoDescricao_GotFocus()
     Call objCT.MeioPagtoDescricao_GotFocus
End Sub

Private Sub MeioPagtoDescricao_KeyPress(KeyAscii As Integer)
     Call objCT.MeioPagtoDescricao_KeyPress(KeyAscii)
End Sub

Private Sub MeioPagtoDescricao_Validate(Cancel As Boolean)
     Call objCT.MeioPagtoDescricao_Validate(Cancel)
End Sub

Private Sub NomePortador_GotFocus()
     Call objCT.NomePortador_GotFocus
End Sub

Private Sub NomePortador_KeyPress(KeyAscii As Integer)
     Call objCT.NomePortador_KeyPress(KeyAscii)
End Sub

Private Sub NomePortador_Validate(Cancel As Boolean)
     Call objCT.NomePortador_Validate(Cancel)
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

Private Sub NumeroMP_GotFocus()
     Call objCT.NumeroMP_GotFocus
End Sub

Private Sub NumeroMP_KeyPress(KeyAscii As Integer)
     Call objCT.NumeroMP_KeyPress(KeyAscii)
End Sub

Private Sub NumeroMP_Validate(Cancel As Boolean)
     Call objCT.NumeroMP_Validate(Cancel)
End Sub

Private Sub NumTitulo_GotFocus()
     Call objCT.NumTitulo_GotFocus
End Sub

Private Sub NumTitulo_KeyPress(KeyAscii As Integer)
     Call objCT.NumTitulo_KeyPress(KeyAscii)
End Sub

Private Sub NumTitulo_Validate(Cancel As Boolean)
     Call objCT.NumTitulo_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Pagamento_Click(Index As Integer)
     Call objCT.Pagamento_Click(Index)
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

Private Sub Portador_Click()
     Call objCT.Portador_Click
End Sub

Private Sub Portador_Validate(Cancel As Boolean)
     Call objCT.Portador_Validate(Cancel)
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

Private Sub SaldoCredito_GotFocus()
     Call objCT.SaldoCredito_GotFocus
End Sub

Private Sub SaldoCredito_KeyPress(KeyAscii As Integer)
     Call objCT.SaldoCredito_KeyPress(KeyAscii)
End Sub

Private Sub SaldoCredito_Validate(Cancel As Boolean)
     Call objCT.SaldoCredito_Validate(Cancel)
End Sub

Private Sub SaldoPA_GotFocus()
     Call objCT.SaldoPA_GotFocus
End Sub

Private Sub SaldoPA_KeyPress(KeyAscii As Integer)
     Call objCT.SaldoPA_KeyPress(KeyAscii)
End Sub

Private Sub SaldoPA_Validate(Cancel As Boolean)
     Call objCT.SaldoPA_Validate(Cancel)
End Sub

Private Sub Selecionada_Click()
     Call objCT.Selecionada_Click
End Sub

Private Sub Selecionado_Click()
     Call objCT.Selecionado_Click
End Sub

Private Sub SelecionadoPA_Click()
     Call objCT.SelecionadoPA_Click
End Sub

Private Sub SiglaDocumento_GotFocus()
     Call objCT.SiglaDocumento_GotFocus
End Sub

Private Sub SiglaDocumento_KeyPress(KeyAscii As Integer)
     Call objCT.SiglaDocumento_KeyPress(KeyAscii)
End Sub

Private Sub SiglaDocumento_Validate(Cancel As Boolean)
     Call objCT.SiglaDocumento_Validate(Cancel)
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

Private Sub TituloFim_Change()
     Call objCT.TituloFim_Change
End Sub

Private Sub TituloFim_GotFocus()
     Call objCT.TituloFim_GotFocus
End Sub

Private Sub TituloFim_Validate(Cancel As Boolean)
     Call objCT.TituloFim_Validate(Cancel)
End Sub

Private Sub TituloInic_Change()
     Call objCT.TituloInic_Change
End Sub

Private Sub TituloInic_GotFocus()
     Call objCT.TituloInic_GotFocus
End Sub

Private Sub UpDownDataBaixa_DownClick()
     Call objCT.UpDownDataBaixa_DownClick
End Sub

Private Sub UpDownDataBaixa_UpClick()
     Call objCT.UpDownDataBaixa_UpClick
End Sub

Private Sub UpDownEmissaoFim_DownClick()
     Call objCT.UpDownEmissaoFim_DownClick
End Sub

Private Sub UpDownEmissaoFim_UpClick()
     Call objCT.UpDownEmissaoFim_UpClick
End Sub

Private Sub UpDownEmissaoInic_DownClick()
     Call objCT.UpDownEmissaoInic_DownClick
End Sub

Private Sub UpDownEmissaoInic_UpClick()
     Call objCT.UpDownEmissaoInic_UpClick
End Sub

Private Sub UpDownVencFim_DownClick()
     Call objCT.UpDownVencFim_DownClick
End Sub

Private Sub UpDownVencFim_UpClick()
     Call objCT.UpDownVencFim_UpClick
End Sub

Private Sub UpDownVencInic_DownClick()
     Call objCT.UpDownVencInic_DownClick
End Sub

Private Sub UpDownVencInic_UpClick()
     Call objCT.UpDownVencInic_UpClick
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTBaixaPag
    Set objCT.objUserControl = Me
End Sub

Private Sub ValorAPagar_GotFocus()
     Call objCT.ValorAPagar_GotFocus
End Sub

Private Sub ValorAPagar_KeyPress(KeyAscii As Integer)
     Call objCT.ValorAPagar_KeyPress(KeyAscii)
End Sub

Private Sub ValorAPagar_Validate(Cancel As Boolean)
     Call objCT.ValorAPagar_Validate(Cancel)
End Sub

Private Sub ValorBaixado_GotFocus()
     Call objCT.ValorBaixado_GotFocus
End Sub

Private Sub ValorBaixado_Change()
     Call objCT.ValorBaixado_Change
End Sub

Private Sub ValorBaixado_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBaixado_KeyPress(KeyAscii)
End Sub

Private Sub ValorBaixado_Validate(Cancel As Boolean)
     Call objCT.ValorBaixado_Validate(Cancel)
End Sub

Private Sub ValorCredito_GotFocus()
     Call objCT.ValorCredito_GotFocus
End Sub

Private Sub ValorCredito_KeyPress(KeyAscii As Integer)
     Call objCT.ValorCredito_KeyPress(KeyAscii)
End Sub

Private Sub ValorCredito_Validate(Cancel As Boolean)
     Call objCT.ValorCredito_Validate(Cancel)
End Sub

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDesconto_GotFocus()
     Call objCT.ValorDesconto_GotFocus
End Sub

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)
     Call objCT.ValorDesconto_KeyPress(KeyAscii)
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub ValorJuros_Change()
     Call objCT.ValorJuros_Change
End Sub

Private Sub ValorMulta_Change()
     Call objCT.ValorMulta_Change
End Sub

Private Sub ValorMulta_GotFocus()
     Call objCT.ValorMulta_GotFocus
End Sub

Private Sub ValorMulta_KeyPress(KeyAscii As Integer)
     Call objCT.ValorMulta_KeyPress(KeyAscii)
End Sub

Private Sub ValorMulta_Validate(Cancel As Boolean)
     Call objCT.ValorMulta_Validate(Cancel)
End Sub

Private Sub ValorJuros_GotFocus()
     Call objCT.ValorJuros_GotFocus
End Sub

Private Sub ValorJuros_KeyPress(KeyAscii As Integer)
     Call objCT.ValorJuros_KeyPress(KeyAscii)
End Sub

Private Sub ValorJuros_Validate(Cancel As Boolean)
     Call objCT.ValorJuros_Validate(Cancel)
End Sub

Private Sub Selecionada_GotFocus()
     Call objCT.Selecionada_GotFocus
End Sub

Private Sub Selecionada_KeyPress(KeyAscii As Integer)
     Call objCT.Selecionada_KeyPress(KeyAscii)
End Sub

Private Sub Selecionada_Validate(Cancel As Boolean)
     Call objCT.Selecionada_Validate(Cancel)
End Sub

Private Sub SelecionadoPA_GotFocus()
     Call objCT.SelecionadoPA_GotFocus
End Sub

Private Sub SelecionadoPA_KeyPress(KeyAscii As Integer)
     Call objCT.SelecionadoPA_KeyPress(KeyAscii)
End Sub

Private Sub SelecionadoPA_Validate(Cancel As Boolean)
     Call objCT.SelecionadoPA_Validate(Cancel)
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

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros(Optional objBaixaPag As ClassBaixaPagar) As Long
     Trata_Parametros = objCT.Trata_Parametros(objBaixaPag)
End Function

Private Sub ValorPA_GotFocus()
     Call objCT.ValorPA_GotFocus
End Sub

Private Sub ValorPA_KeyPress(KeyAscii As Integer)
     Call objCT.ValorPA_KeyPress(KeyAscii)
End Sub

Private Sub ValorPA_Validate(Cancel As Boolean)
     Call objCT.ValorPA_Validate(Cancel)
End Sub

Private Sub ValorPago_Change()
     Call objCT.ValorPago_Change
End Sub

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub VencFim_Change()
     Call objCT.VencFim_Change
End Sub

Private Sub VencFim_GotFocus()
     Call objCT.VencFim_GotFocus
End Sub

Private Sub VencFim_Validate(Cancel As Boolean)
     Call objCT.VencFim_Validate(Cancel)
End Sub

Private Sub VencInic_Change()
     Call objCT.VencInic_Change
End Sub

Private Sub VencInic_GotFocus()
     Call objCT.VencInic_GotFocus
End Sub

Private Sub VencInic_Validate(Cancel As Boolean)
     Call objCT.VencInic_Validate(Cancel)
End Sub

Private Sub CTBBotaoModeloPadrao_Click()
     Call objCT.CTBBotaoModeloPadrao_Click
End Sub

Private Sub CTBModelo_Click()
     Call objCT.CTBModelo_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub CTBGridContabil_EnterCell()
     Call objCT.CTBGridContabil_EnterCell
End Sub

Private Sub CTBGridContabil_GotFocus()
     Call objCT.CTBGridContabil_GotFocus
End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGridContabil_KeyPress(KeyAscii)
End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.CTBGridContabil_KeyDown(KeyCode, Shift)
End Sub

Private Sub CTBGridContabil_LeaveCell()
     Call objCT.CTBGridContabil_LeaveCell
End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)
     Call objCT.CTBGridContabil_Validate(Cancel)
End Sub

Private Sub CTBGridContabil_RowColChange()
     Call objCT.CTBGridContabil_RowColChange
End Sub

Private Sub CTBGridContabil_Scroll()
     Call objCT.CTBGridContabil_Scroll
End Sub

Private Sub CTBConta_Change()
     Call objCT.CTBConta_Change
End Sub

Private Sub CTBConta_GotFocus()
     Call objCT.CTBConta_GotFocus
End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)
     Call objCT.CTBConta_KeyPress(KeyAscii)
End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)
     Call objCT.CTBConta_Validate(Cancel)
End Sub

Private Sub CTBCcl_Change()
     Call objCT.CTBCcl_Change
End Sub

Private Sub CTBCcl_GotFocus()
     Call objCT.CTBCcl_GotFocus
End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCcl_KeyPress(KeyAscii)
End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)
     Call objCT.CTBCcl_Validate(Cancel)
End Sub

Private Sub CTBCredito_Change()
     Call objCT.CTBCredito_Change
End Sub

Private Sub CTBCredito_GotFocus()
     Call objCT.CTBCredito_GotFocus
End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCredito_KeyPress(KeyAscii)
End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)
     Call objCT.CTBCredito_Validate(Cancel)
End Sub

Private Sub CTBDebito_Change()
     Call objCT.CTBDebito_Change
End Sub

Private Sub CTBDebito_GotFocus()
     Call objCT.CTBDebito_GotFocus
End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBDebito_KeyPress(KeyAscii)
End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)
     Call objCT.CTBDebito_Validate(Cancel)
End Sub

Private Sub CTBSeqContraPartida_Change()
     Call objCT.CTBSeqContraPartida_Change
End Sub

Private Sub CTBSeqContraPartida_GotFocus()
     Call objCT.CTBSeqContraPartida_GotFocus
End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.CTBSeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.CTBSeqContraPartida_Validate(Cancel)
End Sub

Private Sub CTBHistorico_Change()
     Call objCT.CTBHistorico_Change
End Sub

Private Sub CTBHistorico_GotFocus()
     Call objCT.CTBHistorico_GotFocus
End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)
     Call objCT.CTBHistorico_KeyPress(KeyAscii)
End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)
     Call objCT.CTBHistorico_Validate(Cancel)
End Sub

Private Sub CTBLancAutomatico_Click()
    Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
    Call objCT.CTBAglutina_Click
End Sub

Private Sub CTBAglutina_GotFocus()
     Call objCT.CTBAglutina_GotFocus
End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)
     Call objCT.CTBAglutina_KeyPress(KeyAscii)
End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)
     Call objCT.CTBAglutina_Validate(Cancel)
End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_NodeClick(Node)
End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_Expand(Node)
End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwCcls_NodeClick(Node)
End Sub

Private Sub CTBListHistoricos_DblClick()
     Call objCT.CTBListHistoricos_DblClick
End Sub

Private Sub CTBBotaoLimparGrid_Click()
     Call objCT.CTBBotaoLimparGrid_Click
End Sub

Private Sub CTBLote_Change()
     Call objCT.CTBLote_Change
End Sub

Private Sub CTBLote_GotFocus()
     Call objCT.CTBLote_GotFocus
End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)
     Call objCT.CTBLote_Validate(Cancel)
End Sub

Private Sub CTBDataContabil_Change()
     Call objCT.CTBDataContabil_Change
End Sub

Private Sub CTBDataContabil_GotFocus()
     Call objCT.CTBDataContabil_GotFocus
End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)
     Call objCT.CTBDataContabil_Validate(Cancel)
End Sub

Private Sub CTBDocumento_Change()
     Call objCT.CTBDocumento_Change
End Sub

Private Sub CTBDocumento_GotFocus()
    Call objCT.CTBDocumento_GotFocus
End Sub
Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub CTBLabelDoc_Click()
     Call objCT.CTBLabelDoc_Click
End Sub

Private Sub CTBLabelLote_Click()
     Call objCT.CTBLabelLote_Click
End Sub

Public Function Form_Load_Ocx() As Object

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



Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub FornecLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecLabel, Source, X, Y)
End Sub

Private Sub FornecLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
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

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
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

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub ValorPago_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPago, Source, X, Y)
End Sub

Private Sub ValorPago_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPago, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub TotalBaixar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalBaixar, Source, X, Y)
End Sub

Private Sub TotalBaixar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalBaixar, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

'##############################################################
'Inserido por Wagner 06/06/2006
Private Sub GridChequePre_Click()
     Call objCT.GridChequePre_Click
End Sub

Private Sub GridChequePre_GotFocus()
     Call objCT.GridChequePre_GotFocus
End Sub

Private Sub GridChequePre_EnterCell()
     Call objCT.GridChequePre_EnterCell
End Sub

Private Sub GridChequePre_LeaveCell()
     Call objCT.GridChequePre_LeaveCell
End Sub

Private Sub GridChequePre_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridChequePre_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridChequePre_KeyPress(KeyAscii As Integer)
     Call objCT.GridChequePre_KeyPress(KeyAscii)
End Sub

Private Sub GridChequePre_Validate(Cancel As Boolean)
     Call objCT.GridChequePre_Validate(Cancel)
End Sub

Private Sub GridChequePre_RowColChange()
     Call objCT.GridChequePre_RowColChange
End Sub

Private Sub GridChequePre_Scroll()
     Call objCT.GridChequePre_Scroll
End Sub

Private Sub SelecionadoCT_Click()
     Call objCT.SelecionadoCT_Click
End Sub

Private Sub SelecionadoCT_GotFocus()
     Call objCT.SelecionadoCT_GotFocus
End Sub

Private Sub SelecionadoCT_KeyPress(KeyAscii As Integer)
     Call objCT.SelecionadoCT_KeyPress(KeyAscii)
End Sub

Private Sub SelecionadoCT_Validate(Cancel As Boolean)
     Call objCT.SelecionadoCT_Validate(Cancel)
End Sub
'#######################################################################

Private Sub CTBGerencial_Click()
    Call objCT.CTBGerencial_Click
End Sub

Private Sub CTBGerencial_GotFocus()
    Call objCT.CTBGerencial_GotFocus
End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)
    Call objCT.CTBGerencial_KeyPress(KeyAscii)
End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)
    Call objCT.CTBGerencial_Validate(Cancel)
End Sub

Private Sub Ordenacao_Change()
    objCT.Ordenacao_Change
End Sub

Private Sub Ordenacao_Click()
    objCT.Ordenacao_Click
End Sub

Private Sub TipoDocTodos_Click()
     Call objCT.TipoDocTodos_Click
End Sub

Private Sub TipoDocApenas_Click()
     Call objCT.TipoDocApenas_Click
End Sub

Private Sub TipoDocSeleciona_Change()
     Call objCT.TipoDocSeleciona_Change
End Sub

Private Sub TipoDocSeleciona_Click()
     Call objCT.TipoDocSeleciona_Change
End Sub

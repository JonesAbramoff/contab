VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl NFiscalFatEntrada 
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5745
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4695
      Index           =   3
      Left            =   210
      TabIndex        =   222
      Top             =   960
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame frame2 
         Caption         =   "Dados do Fornecedor para fins de beneficiamento"
         Height          =   555
         Index           =   6
         Left            =   30
         TabIndex        =   244
         Top             =   4140
         Width           =   9060
         Begin VB.ComboBox FilialFornBenef 
            Height          =   315
            Left            =   6870
            TabIndex        =   210
            Top             =   195
            Width           =   1860
         End
         Begin MSMask.MaskEdBox FornecedorBenef 
            Height          =   315
            Left            =   4125
            TabIndex        =   209
            Top             =   195
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label FornecedorBenefLabel 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3015
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   246
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   49
            Left            =   6360
            TabIndex        =   245
            Top             =   255
            Width           =   465
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "Complemento"
         Height          =   2250
         Index           =   4
         Left            =   30
         TabIndex        =   238
         Top             =   1350
         Width           =   9060
         Begin VB.TextBox Observacao 
            Height          =   315
            Left            =   5715
            MaxLength       =   40
            TabIndex        =   204
            Top             =   1875
            Width           =   3255
         End
         Begin VB.CheckBox MsgAutomatica 
            Caption         =   "Calcula as mensagens automaticamente"
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
            Left            =   1815
            TabIndex        =   199
            Top             =   120
            Value           =   1  'Checked
            Width           =   4755
         End
         Begin VB.TextBox Mensagem 
            Height          =   750
            Left            =   1815
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   201
            Top             =   1110
            Width           =   7155
         End
         Begin VB.TextBox MensagemCorpo 
            Height          =   750
            Left            =   1815
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   200
            Top             =   330
            Width           =   7155
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   315
            Left            =   4080
            TabIndex        =   203
            Top             =   1890
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   315
            Left            =   1815
            TabIndex        =   202
            Top             =   1890
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto:"
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
            Height          =   210
            Index           =   53
            Left            =   780
            TabIndex        =   243
            Top             =   1920
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Peso L�q.:"
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
            Index           =   54
            Left            =   3165
            TabIndex        =   242
            Top             =   1905
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Obs:"
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
            Index           =   11
            Left            =   5265
            TabIndex        =   241
            Top             =   1935
            Width           =   405
         End
         Begin VB.Label MensagemLabel 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem N.Fiscal:"
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
            Left            =   60
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   240
            Top             =   1140
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Msg Corpo da N.F.:"
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
            Index           =   8
            Left            =   90
            TabIndex        =   239
            Top             =   375
            Width           =   1665
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "Nota Fiscal Original"
         Height          =   570
         Index           =   8
         Left            =   30
         TabIndex        =   233
         Top             =   3585
         Width           =   9060
         Begin VB.ComboBox SerieNFiscalOriginal 
            Height          =   315
            Left            =   600
            TabIndex        =   205
            Top             =   210
            Width           =   765
         End
         Begin VB.ComboBox FilialFornNFOrig 
            Height          =   315
            Left            =   6855
            TabIndex        =   208
            Top             =   210
            Width           =   1860
         End
         Begin MSMask.MaskEdBox NFiscalOriginal 
            Height          =   315
            Left            =   2130
            TabIndex        =   206
            Top             =   210
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornNFOrig 
            Height          =   315
            Left            =   4095
            TabIndex        =   207
            Top             =   210
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label NFiscalOriginalLabel 
            Caption         =   "N�mero:"
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
            Left            =   1395
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   237
            Top             =   270
            Width           =   720
         End
         Begin VB.Label SerieOriginalLabel 
            AutoSize        =   -1  'True
            Caption         =   "S�rie:"
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
            Left            =   60
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   236
            Top             =   270
            Width           =   510
         End
         Begin VB.Label LabelFilialFornNFOrig 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6360
            TabIndex        =   235
            Top             =   270
            Width           =   465
         End
         Begin VB.Label LabelFornNFOrig 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3045
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   234
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "Volumes"
         Height          =   480
         Index           =   1
         Left            =   30
         TabIndex        =   228
         Top             =   885
         Width           =   9060
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7140
            MaxLength       =   20
            TabIndex        =   198
            Top             =   135
            Width           =   1815
         End
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   3285
            TabIndex        =   196
            Top             =   135
            Width           =   1335
         End
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   5280
            TabIndex        =   197
            Top             =   135
            Width           =   1335
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1815
            TabIndex        =   195
            Top             =   135
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N� :"
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
            Index           =   57
            Left            =   6750
            TabIndex        =   232
            Top             =   195
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   58
            Left            =   720
            TabIndex        =   231
            Top             =   180
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Esp�cie:"
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
            Index           =   6
            Left            =   2520
            TabIndex        =   230
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
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
            Index           =   60
            Left            =   4680
            TabIndex        =   229
            Top             =   195
            Width           =   600
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "Dados de Transporte"
         Height          =   900
         Index           =   5
         Left            =   30
         TabIndex        =   223
         Top             =   0
         Width           =   9060
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   7140
            TabIndex        =   194
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   5265
            MaxLength       =   10
            TabIndex        =   193
            Top             =   540
            Width           =   1290
         End
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   5265
            TabIndex        =   192
            Top             =   210
            Width           =   2595
         End
         Begin VB.Frame frame2 
            Caption         =   "Frete por conta"
            Height          =   675
            Index           =   9
            Left            =   1575
            TabIndex        =   224
            Top             =   180
            Width           =   1965
            Begin VB.ComboBox TipoFrete 
               Height          =   315
               Left            =   285
               TabIndex        =   247
               Top             =   255
               Width           =   1440
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "U.F.:"
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
            Index           =   69
            Left            =   6645
            TabIndex        =   227
            Top             =   615
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Placa Ve�culo:"
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
            Index           =   0
            Left            =   3945
            TabIndex        =   226
            Top             =   570
            Width           =   1275
         End
         Begin VB.Label TransportadoraLabel 
            AutoSize        =   -1  'True
            Caption         =   "Transportadora:"
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
            Left            =   3855
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   225
            Top             =   270
            Width           =   1365
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   4575
      Index           =   2
      Left            =   180
      TabIndex        =   22
      Top             =   1005
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CommandButton BotaoCcls 
         Caption         =   "Centros de Custo"
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
         Left            =   5895
         TabIndex        =   190
         Top             =   4140
         Width           =   1620
      End
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos"
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
         Left            =   1470
         TabIndex        =   187
         Top             =   4140
         Width           =   1305
      End
      Begin VB.CommandButton BotaoGrade 
         Caption         =   "Grade ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   186
         Top             =   4140
         Width           =   1260
      End
      Begin VB.CommandButton BotaoItemContrato 
         Caption         =   "Item de Contrato"
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
         Left            =   2835
         TabIndex        =   188
         Top             =   4140
         Width           =   1635
      End
      Begin VB.CommandButton BotaoMedicao 
         Caption         =   "Medi��es ..."
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
         Left            =   4545
         TabIndex        =   189
         Top             =   4140
         Width           =   1305
      End
      Begin VB.CommandButton BotaoDocContrato 
         Caption         =   "Contrato ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7605
         TabIndex        =   191
         Top             =   4140
         Width           =   1410
      End
      Begin VB.Frame Frame9 
         Caption         =   "Valores"
         Height          =   1275
         Index           =   1
         Left            =   120
         TabIndex        =   89
         Top             =   2745
         Width           =   8880
         Begin MSMask.MaskEdBox Total 
            Height          =   285
            Left            =   7095
            TabIndex        =   37
            Top             =   900
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   345
            TabIndex        =   33
            Top             =   900
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   -20000
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   885
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   3719
            TabIndex        =   35
            Top             =   915
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   2032
            TabIndex        =   34
            Top             =   900
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   -20000
            TabIndex        =   174
            Top             =   0
            Width           =   1500
         End
         Begin VB.Label SubTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7095
            TabIndex        =   106
            Top             =   390
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Seguro"
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
            Index           =   48
            Left            =   2475
            TabIndex        =   105
            Top             =   705
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
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
            Index           =   19
            Left            =   -20000
            TabIndex        =   104
            Top             =   705
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
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
            Index           =   5
            Left            =   870
            TabIndex        =   103
            Top             =   705
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Despesas"
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
            Index           =   10
            Left            =   4049
            TabIndex        =   102
            Top             =   705
            Width           =   840
         End
         Begin VB.Label LabelTotais 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Left            =   7620
            TabIndex        =   101
            Top             =   690
            Width           =   450
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5406
            TabIndex        =   100
            Top             =   900
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "IPI"
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
            Height          =   165
            Index           =   64
            Left            =   6029
            TabIndex        =   99
            Top             =   705
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Produtos"
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
            Index           =   221
            Left            =   7463
            TabIndex        =   98
            Top             =   195
            Width           =   765
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   345
            TabIndex        =   97
            Top             =   390
            Width           =   1500
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2032
            TabIndex        =   96
            Top             =   390
            Width           =   1500
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3719
            TabIndex        =   95
            Top             =   390
            Width           =   1500
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5406
            TabIndex        =   94
            Top             =   390
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ICMS"
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
            Index           =   2
            Left            =   2550
            TabIndex        =   93
            Top             =   195
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Base ICMS"
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
            Index           =   17
            Left            =   623
            TabIndex        =   92
            Top             =   195
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Base ICMS Subst"
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
            Index           =   217
            Left            =   3727
            TabIndex        =   91
            Top             =   195
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ICMS Subst"
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
            Index           =   219
            Left            =   5654
            TabIndex        =   90
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Itens"
         Height          =   2685
         Index           =   6
         Left            =   165
         TabIndex        =   88
         Top             =   -15
         Width           =   8865
         Begin MSMask.MaskEdBox DataCobranca 
            Height          =   255
            Left            =   6225
            TabIndex        =   215
            Top             =   1770
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Item 
            Height          =   225
            Left            =   5805
            TabIndex        =   184
            Top             =   1770
            Width           =   585
            _ExtentX        =   1032
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
            Format          =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contrato 
            Height          =   225
            Left            =   4590
            TabIndex        =   185
            Top             =   1725
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   105
            MaxLength       =   50
            TabIndex        =   28
            Top             =   825
            Width           =   2385
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            TabIndex        =   24
            Text            =   "UnidadeMed"
            Top             =   300
            Width           =   660
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   225
            Left            =   4200
            TabIndex        =   30
            Top             =   1230
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   5490
            TabIndex        =   31
            Top             =   720
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
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   4245
            TabIndex        =   29
            Top             =   765
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox ValorUnitario 
            Height          =   225
            Left            =   3315
            TabIndex        =   26
            Top             =   315
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2295
            TabIndex        =   25
            Top             =   330
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   300
            TabIndex        =   23
            Top             =   270
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotal 
            Height          =   225
            Left            =   4515
            TabIndex        =   27
            Top             =   315
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1995
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   8685
            _ExtentX        =   15319
            _ExtentY        =   3519
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Index           =   1
      Left            =   150
      TabIndex        =   21
      Top             =   1020
      Width           =   9180
      Begin VB.Frame Frame1 
         Caption         =   "Projetos"
         Height          =   615
         Index           =   10
         Left            =   165
         TabIndex        =   219
         Top             =   2880
         Width           =   8865
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   4485
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   210
            Width           =   2550
         End
         Begin VB.CommandButton BotaoProjetos 
            Caption         =   "..."
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
            Left            =   3150
            TabIndex        =   16
            Top             =   210
            Width           =   495
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   1290
            TabIndex        =   15
            Top             =   225
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Etapa:"
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
            Index           =   4
            Left            =   3900
            TabIndex        =   221
            Top             =   270
            Width           =   570
         End
         Begin VB.Label LabelProjeto 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   555
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   220
            Top             =   270
            Width           =   675
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Complemento"
         Height          =   1005
         Index           =   11
         Left            =   180
         TabIndex        =   180
         Top             =   3510
         Width           =   8850
         Begin VB.ComboBox SubConta 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   4575
         End
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   210
            Width           =   4575
         End
         Begin MSMask.MaskEdBox TaxaConversao 
            Height          =   315
            Left            =   6480
            TabIndex        =   19
            Top             =   195
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "###,##0.00##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "NFe:"
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
            Index           =   65
            Left            =   5970
            TabIndex        =   218
            Top             =   675
            Width           =   420
         End
         Begin VB.Label NumNFe 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6480
            TabIndex        =   217
            Top             =   630
            Width           =   2310
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sub Conta:"
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
            Index           =   26
            Left            =   255
            TabIndex        =   183
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Index           =   25
            Left            =   390
            TabIndex        =   182
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Taxa:"
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
            Index           =   63
            Left            =   5880
            TabIndex        =   181
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "Identifica��o"
         Height          =   1275
         Index           =   0
         Left            =   165
         TabIndex        =   113
         Top             =   0
         Width           =   8865
         Begin VB.CheckBox EletronicaFed 
            Caption         =   "Eletr�nica Federal"
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
            Left            =   6510
            TabIndex        =   7
            Top             =   960
            Width           =   2070
         End
         Begin VB.CommandButton Recebimento 
            Height          =   360
            Left            =   1995
            Picture         =   "NFiscalFatEntradaHic.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Trazer Dados"
            Top             =   135
            Width           =   360
         End
         Begin VB.CommandButton BotaoLimparNF 
            Height          =   315
            Left            =   5685
            Picture         =   "NFiscalFatEntradaHic.ctx":03D2
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Numera��o Autom�tica"
            Top             =   885
            Width           =   345
         End
         Begin VB.ComboBox TipoNFiscal 
            Height          =   315
            ItemData        =   "NFiscalFatEntradaHic.ctx":0904
            Left            =   1320
            List            =   "NFiscalFatEntradaHic.ctx":0906
            TabIndex        =   2
            Top             =   525
            Width           =   4725
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   1305
            TabIndex        =   4
            Top             =   870
            Width           =   765
         End
         Begin MSMask.MaskEdBox NFiscal 
            Height          =   315
            Left            =   4485
            TabIndex        =   5
            Top             =   885
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NaturezaOp 
            Height          =   315
            Left            =   8295
            TabIndex        =   3
            Top             =   570
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
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
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumRecebimento 
            Height          =   315
            Left            =   1320
            TabIndex        =   0
            Top             =   165
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Index           =   1
            Left            =   3840
            TabIndex        =   146
            Top             =   225
            Width           =   615
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4500
            TabIndex        =   145
            Top             =   165
            Width           =   1530
         End
         Begin VB.Label RecebimentoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Recebimento:"
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
            Left            =   75
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   144
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label NFiscalInterna 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4500
            TabIndex        =   143
            Top             =   885
            Width           =   1215
         End
         Begin VB.Label Label1 
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
            Index           =   59
            Left            =   810
            TabIndex        =   117
            Top             =   570
            Width           =   450
         End
         Begin VB.Label NFiscalLabel 
            Caption         =   "N�mero:"
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
            Left            =   3675
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   116
            Top             =   915
            Width           =   720
         End
         Begin VB.Label SerieLabel 
            AutoSize        =   -1  'True
            Caption         =   "S�rie:"
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
            Left            =   750
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   115
            Top             =   915
            Width           =   510
         End
         Begin VB.Label NaturezaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Opera��o:"
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
            Left            =   6510
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   114
            Top             =   630
            Width           =   1725
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dados do Fornecedor"
         Height          =   915
         Index           =   3
         Left            =   165
         TabIndex        =   110
         Top             =   1320
         Width           =   8865
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4485
            TabIndex        =   9
            Top             =   240
            Width           =   2250
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1290
            TabIndex        =   8
            Top             =   225
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label RazaoSocial 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1290
            TabIndex        =   179
            Top             =   570
            Width           =   4410
         End
         Begin VB.Label CNPJ 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6525
            TabIndex        =   178
            Top             =   570
            Width           =   2235
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
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
            Index           =   201
            Left            =   5910
            TabIndex        =   177
            Top             =   615
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Raz�o Social:"
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
            Index           =   200
            Left            =   75
            TabIndex        =   176
            Top             =   600
            Width           =   1200
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
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   112
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   7
            Left            =   3945
            TabIndex        =   111
            Top             =   285
            Width           =   465
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Datas"
         Height          =   600
         Index           =   10
         Left            =   165
         TabIndex        =   107
         Top             =   2250
         Width           =   8865
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   2355
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   195
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownEntrada 
            Height          =   300
            Left            =   5550
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   195
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1275
            TabIndex        =   10
            Top             =   195
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEntrada 
            Height          =   300
            Left            =   4470
            TabIndex        =   12
            Top             =   195
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HoraEntrada 
            Height          =   300
            Left            =   7875
            TabIndex        =   14
            Top             =   195
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora Entrada:"
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
            Index           =   38
            Left            =   6630
            TabIndex        =   171
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Entrada:"
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
            Index           =   47
            Left            =   3210
            TabIndex        =   109
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Emiss�o:"
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
            Index           =   9
            Left            =   450
            TabIndex        =   108
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4575
      Index           =   6
      Left            =   180
      TabIndex        =   47
      Top             =   1005
      Visible         =   0   'False
      Width           =   9180
      Begin TelasESTHic.TabTributacaoFat TabTrib 
         Height          =   4620
         Left            =   135
         TabIndex        =   216
         Top             =   0
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   8149
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4560
      Index           =   5
      Left            =   180
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CheckBox PagamentoAutomatico 
         Caption         =   "Calcula Pagamento Automaticamente"
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
         Left            =   420
         TabIndex        =   39
         Top             =   210
         Value           =   1  'Checked
         Width           =   3705
      End
      Begin VB.Frame frame2 
         Caption         =   "Pagamento"
         Height          =   3720
         Index           =   2
         Left            =   195
         TabIndex        =   86
         Top             =   630
         Width           =   8760
         Begin MSMask.MaskEdBox CodigodeBarras 
            Height          =   315
            Left            =   630
            TabIndex        =   214
            Top             =   2820
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   57
            Mask            =   "#####.#####.#####.######.#####.######.#.#################"
            PromptChar      =   " "
         End
         Begin VB.ComboBox ComboCobrador 
            Height          =   315
            Left            =   5025
            TabIndex        =   173
            Top             =   2220
            Width           =   2295
         End
         Begin VB.ComboBox ComboPortador 
            Height          =   315
            Left            =   2100
            TabIndex        =   172
            Top             =   2250
            Width           =   2445
         End
         Begin VB.CheckBox Suspenso 
            Height          =   225
            Left            =   6630
            TabIndex        =   45
            Top             =   1185
            Width           =   1230
         End
         Begin VB.ComboBox TipoCobranca 
            Height          =   315
            Left            =   4575
            TabIndex        =   44
            Top             =   1155
            Width           =   2010
         End
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   4380
            TabIndex        =   40
            Top             =   345
            Width           =   1815
         End
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   225
            Left            =   2265
            TabIndex        =   42
            Top             =   1155
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   3360
            TabIndex        =   43
            Top             =   1140
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   1140
            TabIndex        =   41
            Top             =   1170
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
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2565
            Left            =   420
            TabIndex        =   46
            Top             =   960
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   4524
            _Version        =   393216
            Rows            =   13
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label CondPagtoLabel 
            Caption         =   "Condi��o de Pagamento:"
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
            Left            =   2160
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   87
            Top             =   390
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   7
      Left            =   180
      TabIndex        =   48
      Top             =   1005
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4440
         TabIndex        =   213
         Tag             =   "1"
         Top             =   2520
         Width           =   870
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
         Left            =   6330
         TabIndex        =   54
         Top             =   345
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
         Left            =   6330
         TabIndex        =   52
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   870
         Width           =   2700
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
         Left            =   7785
         TabIndex        =   53
         Top             =   30
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4920
         TabIndex        =   62
         Top             =   1320
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
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6360
         TabIndex        =   66
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   63
         Top             =   1620
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   64
         Top             =   2010
         Width           =   870
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descri��o do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   121
         Top             =   3450
         Width           =   5895
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   125
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBLabel 
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
            Index           =   7
            Left            =   1125
            TabIndex        =   124
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
            TabIndex        =   123
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   122
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
         Left            =   3480
         TabIndex        =   57
         Top             =   930
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   58
         Top             =   1305
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
         TabIndex        =   61
         Top             =   1335
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
         TabIndex        =   60
         Top             =   1275
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
         TabIndex        =   59
         Top             =   1320
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
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   51
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
         TabIndex        =   50
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
         Left            =   3780
         TabIndex        =   49
         Top             =   135
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
         TabIndex        =   65
         Top             =   1170
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
         Left            =   6360
         TabIndex        =   67
         Top             =   1560
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
         Left            =   6360
         TabIndex        =   68
         Top             =   1560
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
      Begin VB.Label CTBLabel 
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
         Index           =   1
         Left            =   6390
         TabIndex        =   55
         Top             =   660
         Width           =   690
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
         TabIndex        =   142
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
         TabIndex        =   141
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label CTBLabel 
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
         Index           =   8
         Left            =   45
         TabIndex        =   140
         Top             =   570
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   139
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabel 
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
         Index           =   0
         Left            =   45
         TabIndex        =   133
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label CTBLabel 
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
         Index           =   13
         Left            =   1995
         TabIndex        =   132
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   131
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   130
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBLabel 
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
         Index           =   4
         Left            =   4230
         TabIndex        =   129
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   128
         Top             =   120
         Width           =   1530
      End
      Begin VB.Label CTBLabel 
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
         Index           =   21
         Left            =   45
         TabIndex        =   127
         Top             =   165
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame17"
      Height          =   4575
      Index           =   8
      Left            =   180
      TabIndex        =   148
      Top             =   990
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CommandButton BotaoSerie 
         Caption         =   "S�ries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         TabIndex        =   211
         Top             =   4170
         Width           =   1665
      End
      Begin VB.CommandButton BotaoLotes 
         Caption         =   "Lotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7335
         TabIndex        =   159
         Top             =   4170
         Width           =   1665
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rastreamento do Produto"
         Height          =   4050
         Index           =   12
         Left            =   45
         TabIndex        =   149
         Top             =   15
         Width           =   9030
         Begin VB.ComboBox ProdutoRastro 
            Height          =   315
            ItemData        =   "NFiscalFatEntradaHic.ctx":0908
            Left            =   0
            List            =   "NFiscalFatEntradaHic.ctx":0915
            Style           =   2  'Dropdown List
            TabIndex        =   212
            Top             =   0
            Width           =   1740
         End
         Begin VB.ComboBox EscaninhoRastro 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "NFiscalFatEntradaHic.ctx":0934
            Left            =   3870
            List            =   "NFiscalFatEntradaHic.ctx":093E
            Style           =   2  'Dropdown List
            TabIndex        =   160
            Top             =   165
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSMask.MaskEdBox UMRastro 
            Height          =   240
            Left            =   3240
            TabIndex        =   150
            Top             =   210
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ItemNFRastro 
            Height          =   225
            Left            =   300
            TabIndex        =   151
            Top             =   765
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxRastro 
            Height          =   240
            Left            =   1815
            TabIndex        =   152
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRastro 
            Height          =   225
            Left            =   1995
            TabIndex        =   153
            Top             =   405
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSMask.MaskEdBox LoteRastro 
            Height          =   225
            Left            =   2970
            TabIndex        =   154
            Top             =   405
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteDataRastro 
            Height          =   255
            Left            =   5730
            TabIndex        =   155
            Top             =   405
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialOPRastro 
            Height          =   225
            Left            =   4125
            TabIndex        =   156
            Top             =   405
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSMask.MaskEdBox QuantLoteRastro 
            Height          =   225
            Left            =   6885
            TabIndex        =   157
            Top             =   435
            Width           =   1575
            _ExtentX        =   2778
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
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridRastro 
            Height          =   3315
            Left            =   180
            TabIndex        =   158
            Top             =   330
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   5847
            _Version        =   393216
            Rows            =   51
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   4
      Left            =   180
      TabIndex        =   162
      Top             =   1020
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame Frame1 
         Caption         =   "Distribui��o dos Produtos"
         Height          =   3465
         Index           =   0
         Left            =   300
         TabIndex        =   164
         Top             =   330
         Width           =   8370
         Begin VB.ComboBox ProdutoAlmoxDist 
            Height          =   315
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   175
            Top             =   495
            Width           =   1920
         End
         Begin MSMask.MaskEdBox UMDist 
            Height          =   225
            Left            =   4425
            TabIndex        =   165
            Top             =   120
            Visible         =   0   'False
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxDist 
            Height          =   225
            Left            =   3060
            TabIndex        =   166
            Top             =   135
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantDist 
            Height          =   225
            Left            =   6540
            TabIndex        =   167
            Top             =   105
            Width           =   1470
            _ExtentX        =   2593
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
         Begin MSMask.MaskEdBox ItemNFDist 
            Height          =   225
            Left            =   960
            TabIndex        =   168
            Top             =   105
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDist 
            Height          =   2910
            Left            =   360
            TabIndex        =   169
            Top             =   345
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   5133
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox QuantItemNFDist 
            Height          =   225
            Left            =   5025
            TabIndex        =   170
            Top             =   150
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
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
      Begin VB.CommandButton BotaoLocalizacaoDist 
         Caption         =   "Estoque"
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
         Left            =   6960
         TabIndex        =   163
         Top             =   4140
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   525
      Left            =   6585
      ScaleHeight     =   465
      ScaleWidth      =   2745
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   90
      Width           =   2805
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   1357
         Picture         =   "NFiscalFatEntradaHic.ctx":095A
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoConsultaTitPag 
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
         Left            =   60
         Picture         =   "NFiscalFatEntradaHic.ctx":0AE4
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Consulta de T�tulo � Pagar"
         Top             =   75
         Width           =   765
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   896
         Picture         =   "NFiscalFatEntradaHic.ctx":1366
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   1818
         Picture         =   "NFiscalFatEntradaHic.ctx":14C0
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   2280
         Picture         =   "NFiscalFatEntradaHic.ctx":19F2
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   9
      Left            =   180
      TabIndex        =   69
      Top             =   1005
      Visible         =   0   'False
      Width           =   9180
      Begin VB.ComboBox UnidadeMedBenef 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2790
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   525
         Width           =   645
      End
      Begin VB.TextBox DescricaoItemBenef 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   4710
         MaxLength       =   50
         TabIndex        =   73
         Top             =   570
         Width           =   2640
      End
      Begin VB.CommandButton BotaoProdutosBenef 
         Caption         =   "Produtos"
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
         Left            =   4500
         TabIndex        =   78
         Top             =   4125
         Width           =   1290
      End
      Begin VB.CommandButton BotaoEstoqueBenef 
         Caption         =   "Estoque"
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
         Left            =   5895
         TabIndex        =   79
         Top             =   4125
         Width           =   1335
      End
      Begin VB.CommandButton BotaoPlanoConta 
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
         Height          =   315
         Left            =   7350
         TabIndex        =   80
         Top             =   4125
         Width           =   1815
      End
      Begin MSMask.MaskEdBox ContaContabilProducao 
         Height          =   270
         Left            =   4260
         TabIndex        =   75
         Top             =   1290
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox ContaContabilEst 
         Height          =   240
         Left            =   6495
         TabIndex        =   76
         Top             =   1290
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox AlmoxarifadoBenef 
         Height          =   240
         Left            =   4215
         TabIndex        =   74
         Top             =   930
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantidadeBenef 
         Height          =   240
         Left            =   3615
         TabIndex        =   72
         Top             =   585
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   423
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoBenef 
         Height          =   240
         Left            =   1020
         TabIndex        =   70
         Top             =   555
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridMovimentos 
         Height          =   3510
         Left            =   240
         TabIndex        =   77
         Top             =   435
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6191
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produtos nossos consumidos no processo de fabrica��o dos materias desta nota"
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
         Left            =   285
         TabIndex        =   120
         Top             =   210
         Width           =   6870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Dispon�vel:"
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
         Index           =   61
         Left            =   345
         TabIndex        =   119
         Top             =   4155
         Width           =   2025
      End
      Begin VB.Label QuantDisponivelBenef 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2445
         TabIndex        =   118
         Top             =   4110
         Width           =   1425
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5085
      Left            =   75
      TabIndex        =   85
      Top             =   615
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8969
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribui��o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagamento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributa��o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabiliza��o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rastro"
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
Attribute VB_Name = "NFiscalFatEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTNFiscalFatEntrada
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoExcluir_Click()
    Call objCT.BotaoExcluir_Click
End Sub

'Rastreamento
Private Sub BotaoLotes_Click()
    Call objCT.BotaoLotes_Click
End Sub

Private Sub ItemNFRastro_Change()
     Call objCT.ItemNFRastro_Change
End Sub

Private Sub ItemNFRastro_GotFocus()
     Call objCT.ItemNFRastro_GotFocus
End Sub

Private Sub ItemNFRastro_KeyPress(KeyAscii As Integer)
     Call objCT.ItemNFRastro_KeyPress(KeyAscii)
End Sub

Private Sub ItemNFRastro_Validate(Cancel As Boolean)
     Call objCT.ItemNFRastro_Validate(Cancel)
End Sub

'distribuicao
Private Sub AlmoxRastro_Change()
     Call objCT.AlmoxRastro_Change
End Sub

'distribuicao
Private Sub AlmoxRastro_GotFocus()
     Call objCT.AlmoxRastro_GotFocus
End Sub

'distribuicao
Private Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
     Call objCT.AlmoxRastro_KeyPress(KeyAscii)
End Sub

'distribuicao
Private Sub AlmoxRastro_Validate(Cancel As Boolean)
     Call objCT.AlmoxRastro_Validate(Cancel)
End Sub

Private Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridRastro_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridRastro_Click()
     Call objCT.GridRastro_Click
End Sub

Private Sub GridRastro_EnterCell()
     Call objCT.GridRastro_EnterCell
End Sub

Private Sub GridRastro_GotFocus()
     Call objCT.GridRastro_GotFocus
End Sub

Private Sub GridRastro_KeyPress(KeyAscii As Integer)
     Call objCT.GridRastro_KeyPress(KeyAscii)
End Sub

Private Sub GridRastro_LeaveCell()
     Call objCT.GridRastro_LeaveCell
End Sub

Private Sub GridRastro_Validate(Cancel As Boolean)
     Call objCT.GridRastro_Validate(Cancel)
End Sub

Private Sub GridRastro_Scroll()
     Call objCT.GridRastro_Scroll
End Sub

Private Sub GridRastro_RowColChange()
     Call objCT.GridRastro_RowColChange
End Sub

Private Sub LoteRastro_Change()
     Call objCT.LoteRastro_Change
End Sub

Private Sub LoteRastro_GotFocus()
     Call objCT.LoteRastro_GotFocus
End Sub

Private Sub LoteRastro_KeyPress(KeyAscii As Integer)
     Call objCT.LoteRastro_KeyPress(KeyAscii)
End Sub

Private Sub LoteRastro_Validate(Cancel As Boolean)
     Call objCT.LoteRastro_Validate(Cancel)
End Sub

Private Sub FilialOPRastro_Change()
     Call objCT.FilialOPRastro_Change
End Sub

Private Sub FilialOPRastro_GotFocus()
     Call objCT.FilialOPRastro_GotFocus
End Sub

Private Sub FilialOPRastro_KeyPress(KeyAscii As Integer)
     Call objCT.FilialOPRastro_KeyPress(KeyAscii)
End Sub

Private Sub FilialOPRastro_Validate(Cancel As Boolean)
     Call objCT.FilialOPRastro_Validate(Cancel)
End Sub

Private Sub QuantLoteRastro_Change()
     Call objCT.QuantLoteRastro_Change
End Sub

Private Sub QuantLoteRastro_GotFocus()
     Call objCT.QuantLoteRastro_GotFocus
End Sub

Private Sub QuantLoteRastro_KeyPress(KeyAscii As Integer)
     Call objCT.QuantLoteRastro_KeyPress(KeyAscii)
End Sub

Private Sub QuantLoteRastro_Validate(Cancel As Boolean)
     Call objCT.QuantLoteRastro_Validate(Cancel)
End Sub
'Fim Rastreamento

Private Sub BotaoConsultaTitPag_Click()
    Call objCT.BotaoConsultaTitPag_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTNFiscalFatEntrada
    Set objCT.objUserControl = Me
    'hicare
    Set objCT.gobjInfoUsu = New CTNFFatEntComVGHic
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTNFFatEntComHic
End Sub

Private Sub BotaoLimparNF_Click()
     Call objCT.BotaoLimparNF_Click
End Sub

Private Sub Ccl_Change()
     Call objCT.Ccl_Change
End Sub

Private Sub Ccl_GotFocus()
     Call objCT.Ccl_GotFocus
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)
     Call objCT.Ccl_KeyPress(KeyAscii)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
     Call objCT.Ccl_Validate(Cancel)
End Sub

Private Sub BotaoCcls_Click()
     Call objCT.BotaoCcls_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEntrada_GotFocus()
     Call objCT.DataEntrada_GotFocus
End Sub

Private Sub TipoFrete_Click()
     Call objCT.TipoFrete_Click
End Sub

Private Sub FilialFornBenef_Change()
     Call objCT.FilialFornBenef_Change
End Sub

Private Sub FornecedorBenefLabel_Click()
     Call objCT.FornecedorBenefLabel_Click
End Sub

Private Sub NaturezaOp_GotFocus()
     Call objCT.NaturezaOp_GotFocus
End Sub


Private Sub NFiscal_GotFocus()
     Call objCT.NFiscal_GotFocus
End Sub

Private Sub NFiscalOriginal_GotFocus()
     Call objCT.NFiscalOriginal_GotFocus
End Sub

Private Sub NumRecebimento_GotFocus()
    Call objCT.NumRecebimento_GotFocus
End Sub

Private Sub Serie_Click()
     Call objCT.Serie_Click
End Sub

Private Sub Total_Change()
     Call objCT.Total_Change
End Sub

Private Sub Total_Validate(Cancel As Boolean)
     Call objCT.Total_Validate(Cancel)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub NaturezaLabel_Click()
     Call objCT.NaturezaLabel_Click
End Sub

Private Sub FornecedorLabel_Click()
     Call objCT.FornecedorLabel_Click
End Sub

Private Sub SerieLabel_Click()
     Call objCT.SerieLabel_Click
End Sub

Private Sub NFiscalLabel_Click()
     Call objCT.NFiscalLabel_Click
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub SerieOriginalLabel_Click()
     Call objCT.SerieOriginalLabel_Click
End Sub

Private Sub NFiscalOriginalLabel_Click()
     Call objCT.NFiscalOriginalLabel_Click
End Sub

Private Sub CondPagtoLabel_Click()
     Call objCT.CondPagtoLabel_Click
End Sub

Private Sub RecebimentoLabel_Click()
     Call objCT.BotaoRecebimentos_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long
     Trata_Parametros = objCT.Trata_Parametros(objNFiscal)
End Function

Private Sub TipoNFiscal_Click()
     Call objCT.TipoNFiscal_Click
End Sub

Private Sub TipoNFiscal_Validate(Cancel As Boolean)
     Call objCT.TipoNFiscal_Validate(Cancel)
End Sub

Private Sub NaturezaOp_Validate(Cancel As Boolean)
     Call objCT.NaturezaOp_Validate(Cancel)
End Sub

Private Sub Fornecedor_Change()
     Call objCT.Fornecedor_Change
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
     Call objCT.Fornecedor_Validate(Cancel)
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub DataEntrada_Validate(Cancel As Boolean)
     Call objCT.DataEntrada_Validate(Cancel)
End Sub

Private Sub UpDownEntrada_DownClick()
     Call objCT.UpDownEntrada_DownClick
End Sub

Private Sub UpDownEntrada_UpClick()
     Call objCT.UpDownEntrada_UpClick
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub Serie_Validate(Cancel As Boolean)
     Call objCT.Serie_Validate(Cancel)
End Sub

Private Sub Recebimento_Click()
     Call objCT.Recebimento_Click
End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)
     Call objCT.ValorFrete_Validate(Cancel)
End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)
     Call objCT.ValorSeguro_Validate(Cancel)
End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)
     Call objCT.ValorDespesas_Validate(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub PlacaUF_Validate(Cancel As Boolean)
     Call objCT.PlacaUF_Validate(Cancel)
End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)
     Call objCT.PesoLiquido_Validate(Cancel)
End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)
     Call objCT.PesoBruto_Validate(Cancel)
End Sub

Private Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)
     Call objCT.SerieNFiscalOriginal_Validate(Cancel)
End Sub

Private Sub PagamentoAutomatico_Click()
     Call objCT.PagamentoAutomatico_Click
End Sub

Private Sub CondicaoPagamento_Click()
     Call objCT.CondicaoPagamento_Click
End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)
     Call objCT.CondicaoPagamento_Validate(Cancel)
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

'mario distribuicao
'Private Sub Almoxarifado_Change()
'     Call objCT.Almoxarifado_Change
'End Sub

Private Sub CondicaoPagamento_Change()
     Call objCT.CondicaoPagamento_Change
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
End Sub

Private Sub DataEntrada_Change()
     Call objCT.DataEntrada_Change
End Sub

'horaentrada
Private Sub HoraEntrada_Change()
     Call objCT.HoraEntrada_Change
End Sub

'horaentrada
Private Sub HoraEntrada_Validate(Cancel As Boolean)
     Call objCT.HoraEntrada_Validate(Cancel)
End Sub

'horaentrada
Private Sub HoraEntrada_GotFocus()
     Call objCT.HoraEntrada_GotFocus
End Sub

Private Sub DataVencimentoReal_Change()
     Call objCT.DataVencimentoReal_Change
End Sub

Private Sub Desconto_Change()
     Call objCT.Desconto_Change
End Sub

Private Sub DescricaoItem_Change()
     Call objCT.DescricaoItem_Change
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Mensagem_Change()
     Call objCT.Mensagem_Change
End Sub

Private Sub Observacao_Change()
     Call objCT.Observacao_Change
End Sub

Private Sub NaturezaOp_Change()
     Call objCT.NaturezaOp_Change
End Sub

Private Sub NFiscal_Change()
     Call objCT.NFiscal_Change
End Sub

Private Sub NFiscalOriginal_Change()
     Call objCT.NFiscalOriginal_Change
End Sub

Private Sub PercentDesc_Change()
     Call objCT.PercentDesc_Change
End Sub

Private Sub PesoBruto_Change()
     Call objCT.PesoBruto_Change
End Sub

Private Sub PesoLiquido_Change()
     Call objCT.PesoLiquido_Change
End Sub

Private Sub Placa_Change()
     Call objCT.Placa_Change
End Sub

Private Sub PlacaUF_Change()
     Call objCT.PlacaUF_Change
End Sub

Private Sub PlacaUF_Click()
     Call objCT.PlacaUF_Click
End Sub

Private Sub ValorTotal_Change()
     Call objCT.ValorTotal_Change
End Sub

Private Sub ValorUnitario_Change()
     Call objCT.ValorUnitario_Change
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Produto_GotFocus()
     Call objCT.Produto_GotFocus
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
     Call objCT.Produto_KeyPress(KeyAscii)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub Quantidade_Change()
     Call objCT.Quantidade_Change
End Sub

Private Sub Serie_Change()
     Call objCT.Serie_Change
End Sub

Private Sub SerieNFiscalOriginal_Change()
     Call objCT.SerieNFiscalOriginal_Change
End Sub

Private Sub SerieNFiscalOriginal_Click()
     Call objCT.SerieNFiscalOriginal_Click
End Sub

Private Sub Suspenso_Click()
     Call objCT.Suspenso_Click
End Sub

Private Sub Suspenso_GotFocus()
     Call objCT.Suspenso_GotFocus
End Sub

Private Sub Suspenso_KeyPress(KeyAscii As Integer)
     Call objCT.Suspenso_KeyPress(KeyAscii)
End Sub

Private Sub Suspenso_Validate(Cancel As Boolean)
     Call objCT.Suspenso_Validate(Cancel)
End Sub

Private Sub TipoCobranca_Change()
     Call objCT.TipoCobranca_Change
End Sub

Private Sub TipoCobranca_GotFocus()
     Call objCT.TipoCobranca_GotFocus
End Sub

Private Sub TipoCobranca_KeyPress(KeyAscii As Integer)
     Call objCT.TipoCobranca_KeyPress(KeyAscii)
End Sub

Private Sub TipoCobranca_Validate(Cancel As Boolean)
     Call objCT.TipoCobranca_Validate(Cancel)
End Sub

Private Sub TipoNFiscal_Change()
     Call objCT.TipoNFiscal_Change
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub Transportadora_Click()
     Call objCT.Transportadora_Click
End Sub

Private Sub UnidadeMed_Change()
     Call objCT.UnidadeMed_Change
End Sub

Private Sub UnidadeMed_Click()
     Call objCT.UnidadeMed_Click
End Sub

Private Sub UnidadeMed_GotFocus()
     Call objCT.UnidadeMed_GotFocus
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
     Call objCT.UnidadeMed_KeyPress(KeyAscii)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)
     Call objCT.UnidadeMed_Validate(Cancel)
End Sub

Private Sub ValorDespesas_Change()
     Call objCT.ValorDespesas_Change
End Sub

Private Sub ValorFrete_Change()
     Call objCT.ValorFrete_Change
End Sub

Private Sub ValorParcela_Change()
     Call objCT.ValorParcela_Change
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

Private Sub ValorSeguro_Change()
     Call objCT.ValorSeguro_Change
End Sub

Private Sub VolumeEspecie_Change()
     Call objCT.VolumeEspecie_Change
End Sub

'Inclu�do por Luiz Nogueira em 21/08/03
Private Sub VolumeEspecie_Click()
     Call objCT.VolumeEspecie_Click
End Sub

'Inclu�do por Luiz Nogueira em 21/08/03
Private Sub VolumeEspecie_Validate(Cancel As Boolean)
    Call objCT.VolumeEspecie_Validate(Cancel)
End Sub

Private Sub VolumeMarca_Change()
     Call objCT.VolumeMarca_Change
End Sub

'Inclu�do por Luiz Nogueira em 21/08/03
Private Sub VolumeMarca_Click()
     Call objCT.VolumeMarca_Click
End Sub

'Inclu�do por Luiz Nogueira em 21/08/03
Private Sub VolumeMarca_Validate(Cancel As Boolean)
    Call objCT.VolumeMarca_Validate(Cancel)
End Sub

Private Sub VolumeNumero_Change()
     Call objCT.VolumeNumero_Change
End Sub

Private Sub VolumeQuant_Change()
     Call objCT.VolumeQuant_Change
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

Private Sub GridItens_Click()
     Call objCT.GridItens_Click
End Sub

Private Sub GridItens_EnterCell()
     Call objCT.GridItens_EnterCell
End Sub

Private Sub GridItens_GotFocus()
     Call objCT.GridItens_GotFocus
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)
     Call objCT.GridItens_KeyPress(KeyAscii)
End Sub

Private Sub GridItens_LeaveCell()
     Call objCT.GridItens_LeaveCell
End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
     Call objCT.GridItens_Validate(Cancel)
End Sub

Private Sub GridItens_RowColChange()
     Call objCT.GridItens_RowColChange
End Sub

Private Sub GridItens_Scroll()
     Call objCT.GridItens_Scroll
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridItens_KeyDown(KeyCode, Shift)
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
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

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
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

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_Expand(Node)
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

Private Sub ValorUnitario_GotFocus()
     Call objCT.ValorUnitario_GotFocus
End Sub

Private Sub ValorUnitario_KeyPress(KeyAscii As Integer)
     Call objCT.ValorUnitario_KeyPress(KeyAscii)
End Sub

Private Sub ValorUnitario_Validate(Cancel As Boolean)
     Call objCT.ValorUnitario_Validate(Cancel)
End Sub

Private Sub DescricaoItem_GotFocus()
     Call objCT.DescricaoItem_GotFocus
End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)
     Call objCT.DescricaoItem_KeyPress(KeyAscii)
End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)
     Call objCT.DescricaoItem_Validate(Cancel)
End Sub

Private Sub PercentDesc_GotFocus()
     Call objCT.PercentDesc_GotFocus
End Sub

Private Sub PercentDesc_KeyPress(KeyAscii As Integer)
     Call objCT.PercentDesc_KeyPress(KeyAscii)
End Sub

Private Sub PercentDesc_Validate(Cancel As Boolean)
     Call objCT.PercentDesc_Validate(Cancel)
End Sub

Private Sub Desconto_GotFocus()
     Call objCT.Desconto_GotFocus
End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto_KeyPress(KeyAscii)
End Sub

Private Sub Desconto_Validate(Cancel As Boolean)
     Call objCT.Desconto_Validate(Cancel)
End Sub

'mario distribuicao
'Private Sub Almoxarifado_GotFocus()
'     Call objCT.Almoxarifado_GotFocus
'End Sub
'
'Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)
'     Call objCT.Almoxarifado_KeyPress(KeyAscii)
'End Sub
'
'Private Sub Almoxarifado_Validate(Cancel As Boolean)
'     Call objCT.Almoxarifado_Validate(Cancel)
'End Sub

Private Sub Quantidade_GotFocus()
     Call objCT.Quantidade_GotFocus
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
     Call objCT.Quantidade_KeyPress(KeyAscii)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub BotaoPlanoConta_Click()
     Call objCT.BotaoPlanoConta_Click
End Sub

Private Sub ContaContabilEst_Change()
     Call objCT.ContaContabilEst_Change
End Sub

Private Sub ContaContabilEst_GotFocus()
     Call objCT.ContaContabilEst_GotFocus
End Sub

Private Sub ContaContabilEst_KeyPress(KeyAscii As Integer)
     Call objCT.ContaContabilEst_KeyPress(KeyAscii)
End Sub

Private Sub ContaContabilEst_Validate(Cancel As Boolean)
     Call objCT.ContaContabilEst_Validate(Cancel)
End Sub

Private Sub ContaContabilProducao_Change()
     Call objCT.ContaContabilProducao_Change
End Sub

Private Sub ContaContabilProducao_GotFocus()
     Call objCT.ContaContabilProducao_GotFocus
End Sub

Private Sub ContaContabilProducao_KeyPress(KeyAscii As Integer)
     Call objCT.ContaContabilProducao_KeyPress(KeyAscii)
End Sub

Private Sub ContaContabilProducao_Validate(Cancel As Boolean)
     Call objCT.ContaContabilProducao_Validate(Cancel)
End Sub

Private Sub BotaoProdutosBenef_Click()
     Call objCT.BotaoProdutosBenef_Click
End Sub

Private Sub BotaoEstoqueBenef_Click()
     Call objCT.BotaoEstoqueBenef_Click
End Sub

Private Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridMovimentos_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridMovimentos_Click()
     Call objCT.GridMovimentos_Click
End Sub

Private Sub GridMovimentos_EnterCell()
     Call objCT.GridMovimentos_EnterCell
End Sub

Private Sub GridMovimentos_GotFocus()
     Call objCT.GridMovimentos_GotFocus
End Sub

Private Sub GridMovimentos_KeyPress(KeyAscii As Integer)
     Call objCT.GridMovimentos_KeyPress(KeyAscii)
End Sub

Private Sub GridMovimentos_LeaveCell()
     Call objCT.GridMovimentos_LeaveCell
End Sub

Private Sub GridMovimentos_Validate(Cancel As Boolean)
     Call objCT.GridMovimentos_Validate(Cancel)
End Sub

Private Sub GridMovimentos_Scroll()
     Call objCT.GridMovimentos_Scroll
End Sub

Private Sub GridMovimentos_RowColChange()
     Call objCT.GridMovimentos_RowColChange
End Sub

Private Sub AlmoxarifadoBenef_GotFocus()
     Call objCT.AlmoxarifadoBenef_GotFocus
End Sub

Private Sub AlmoxarifadoBenef_KeyPress(KeyAscii As Integer)
     Call objCT.AlmoxarifadoBenef_KeyPress(KeyAscii)
End Sub

Private Sub AlmoxarifadoBenef_Validate(Cancel As Boolean)
     Call objCT.AlmoxarifadoBenef_Validate(Cancel)
End Sub

Private Sub ProdutoBenef_GotFocus()
     Call objCT.ProdutoBenef_GotFocus
End Sub

Private Sub ProdutoBenef_KeyPress(KeyAscii As Integer)
     Call objCT.ProdutoBenef_KeyPress(KeyAscii)
End Sub

Private Sub ProdutoBenef_Validate(Cancel As Boolean)
     Call objCT.ProdutoBenef_Validate(Cancel)
End Sub

Private Sub QuantidadeBenef_GotFocus()
     Call objCT.QuantidadeBenef_GotFocus
End Sub

Private Sub QuantidadeBenef_KeyPress(KeyAscii As Integer)
     Call objCT.QuantidadeBenef_KeyPress(KeyAscii)
End Sub

Private Sub QuantidadeBenef_Validate(Cancel As Boolean)
     Call objCT.QuantidadeBenef_Validate(Cancel)
End Sub

Private Sub UnidadeMedBenef_GotFocus()
     Call objCT.UnidadeMedBenef_GotFocus
End Sub

Private Sub UnidadeMedBenef_KeyPress(KeyAscii As Integer)
     Call objCT.UnidadeMedBenef_KeyPress(KeyAscii)
End Sub

Private Sub UnidadeMedBenef_Validate(Cancel As Boolean)
     Call objCT.UnidadeMedBenef_Validate(Cancel)
End Sub

Private Sub FornecedorBenef_Change()
     Call objCT.FornecedorBenef_Change
End Sub

Private Sub FornecedorBenef_Validate(Cancel As Boolean)
     Call objCT.FornecedorBenef_Validate(Cancel)
End Sub

Private Sub FilialFornBenef_Validate(Cancel As Boolean)
     Call objCT.FilialFornBenef_Validate(Cancel)
End Sub

Private Sub VolumeQuant_GotFocus()
     Call objCT.VolumeQuant_GotFocus
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


Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

'Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label8(Index), Source, X, Y)
'End Sub
'
'Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
'End Sub

'Private Sub Label18_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label18(Index), Source, X, Y)
'End Sub
'
'Private Sub Label18_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label18(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label16_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label16(Index), Source, X, Y)
'End Sub
'
'Private Sub Label16_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label16(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label17_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label17(Index), Source, X, Y)
'End Sub
'
'Private Sub Label17_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label17(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label19_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label19(Index), Source, X, Y)
'End Sub
'
'Private Sub Label19_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label19(Index), Button, Shift, X, Y)
'End Sub

'Private Sub Label20_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label20(Index), Source, X, Y)
'End Sub
'
'Private Sub Label20_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label20(Index), Button, Shift, X, Y)
'End Sub

'Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label2(Index), Source, X, Y)
'End Sub
'
'Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label7_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label7(Index), Source, X, Y)
'End Sub
'
'Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label7(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label30_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label30(Index), Source, X, Y)
'End Sub
'
'Private Sub Label30_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label30(Index), Button, Shift, X, Y)
'End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub RecebimentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RecebimentoLabel, Source, X, Y)
End Sub

Private Sub RecebimentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RecebimentoLabel, Button, Shift, X, Y)
End Sub

Private Sub NFiscalInterna_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalInterna, Source, X, Y)
End Sub

Private Sub NFiscalInterna_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalInterna, Button, Shift, X, Y)
End Sub

Private Sub NFiscalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalLabel, Source, X, Y)
End Sub

Private Sub NFiscalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalLabel, Button, Shift, X, Y)
End Sub

Private Sub SerieLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieLabel, Source, X, Y)
End Sub

Private Sub SerieLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieLabel, Button, Shift, X, Y)
End Sub

Private Sub NaturezaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NaturezaLabel, Source, X, Y)
End Sub

Private Sub NaturezaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NaturezaLabel, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub SubTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SubTotal, Source, X, Y)
End Sub

Private Sub SubTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SubTotal, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub IPIValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor1, Source, X, Y)
End Sub

Private Sub IPIValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSBase1, Source, X, Y)
End Sub

Private Sub ICMSBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSBase1, Button, Shift, X, Y)
End Sub

Private Sub ICMSValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSValor1, Source, X, Y)
End Sub

Private Sub ICMSValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstBase1, Source, X, Y)
End Sub

Private Sub ICMSSubstBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstBase1, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstValor1, Source, X, Y)
End Sub

Private Sub ICMSSubstValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstValor1, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub SerieOriginalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieOriginalLabel, Source, X, Y)
End Sub

Private Sub SerieOriginalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieOriginalLabel, Button, Shift, X, Y)
End Sub

Private Sub NFiscalOriginalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalOriginalLabel, Source, X, Y)
End Sub

Private Sub NFiscalOriginalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalOriginalLabel, Button, Shift, X, Y)
End Sub

Private Sub FornecedorBenefLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorBenefLabel, Source, X, Y)
End Sub

Private Sub FornecedorBenefLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorBenefLabel, Button, Shift, X, Y)
End Sub

Private Sub QuantDisponivelBenef_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDisponivelBenef, Source, X, Y)
End Sub

Private Sub QuantDisponivelBenef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivelBenef, Button, Shift, X, Y)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

'distribuicao inicio
Private Sub GridDist_Click()
     Call objCT.GridDist_Click
End Sub

Private Sub GridDist_EnterCell()
     Call objCT.GridDist_EnterCell
End Sub

Private Sub GridDist_GotFocus()
     Call objCT.GridDist_GotFocus
End Sub

Private Sub GridDist_KeyPress(KeyAscii As Integer)
     Call objCT.GridDist_KeyPress(KeyAscii)
End Sub

Private Sub GridDist_LeaveCell()
     Call objCT.GridDist_LeaveCell
End Sub

Private Sub GridDist_Validate(Cancel As Boolean)
     Call objCT.GridDist_Validate(Cancel)
End Sub

Private Sub GridDist_RowColChange()
     Call objCT.GridDist_RowColChange
End Sub

Private Sub GridDist_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridDist_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridDist_Scroll()
     Call objCT.GridDist_Scroll
End Sub

Public Sub ItemNFDist_Change()

    Call objCT.ItemNFDist_Change

End Sub

Public Sub ItemNFDist_GotFocus()

    Call objCT.ItemNFDist_GotFocus

End Sub

Public Sub ItemNFDist_KeyPress(KeyAscii As Integer)

    Call objCT.ItemNFDist_KeyPress(KeyAscii)

End Sub

Public Sub ItemNFDist_Validate(Cancel As Boolean)

    Call objCT.ItemNFDist_Validate(Cancel)

End Sub

Public Sub AlmoxDist_Change()

    Call objCT.AlmoxDist_Change

End Sub

Public Sub AlmoxDist_GotFocus()

    Call objCT.AlmoxDist_GotFocus

End Sub

Public Sub AlmoxDist_KeyPress(KeyAscii As Integer)

    Call objCT.AlmoxDist_KeyPress(KeyAscii)

End Sub

Public Sub AlmoxDist_Validate(Cancel As Boolean)

    Call objCT.AlmoxDist_Validate(Cancel)

End Sub

Public Sub QuantDist_Change()

    Call objCT.QuantDist_Change

End Sub

Public Sub QuantDist_GotFocus()

    Call objCT.QuantDist_GotFocus

End Sub

Public Sub QuantDist_KeyPress(KeyAscii As Integer)

    Call objCT.QuantDist_KeyPress(KeyAscii)

End Sub

Public Sub QuantDist_Validate(Cancel As Boolean)

    Call objCT.QuantDist_Validate(Cancel)

End Sub

Private Sub BotaoLocalizacaoDist_Click()
     Call objCT.BotaoLocalizacaoDist_Click
End Sub

'fim mario distribuicao

'inicio daniel em 03/10/2001

Private Sub ComboPortador_Click()
     Call objCT.ComboPortador_Click
End Sub

Private Sub ComboPortador_Change()
     Call objCT.ComboPortador_Change
End Sub

Private Sub ComboPortador_GotFocus()
    Call objCT.ComboPortador_GotFocus
End Sub

Private Sub ComboPortador_KeyPress(KeyAscii As Integer)
    Call objCT.ComboPortador_KeyPress(KeyAscii)
End Sub

Private Sub ComboPortador_Validate(Cancel As Boolean)
    Call objCT.ComboPortador_Validate(Cancel)
End Sub

Private Sub ComboCobrador_Click()
     Call objCT.ComboCobrador_Click
End Sub

Private Sub ComboCobrador_Change()
     Call objCT.ComboCobrador_Change
End Sub

Private Sub ComboCobrador_GotFocus()
    Call objCT.ComboCobrador_GotFocus
End Sub

Private Sub ComboCobrador_KeyPress(KeyAscii As Integer)
    Call objCT.ComboCobrador_KeyPress(KeyAscii)
End Sub

Private Sub ComboCobrador_Validate(Cancel As Boolean)
    Call objCT.ComboCobrador_Validate(Cancel)
End Sub

'fim daniel em 03/10/2001

'In�cio Luiz em 25/01/02
Private Sub FilialFornNFOrig_Change()
     Call objCT.FilialFornNFOrig_Change
End Sub

Private Sub FilialFornNFOrig_Validate(Cancel As Boolean)
     Call objCT.FilialFornNFOrig_Validate(Cancel)
End Sub

Private Sub FornNFOrig_Change()
     Call objCT.FornNFOrig_Change
End Sub

Private Sub FornNFOrig_Validate(Cancel As Boolean)
     Call objCT.FornNFOrig_Validate(Cancel)
End Sub

Private Sub LabelFornNFOrig_Click()
     Call objCT.LabelFornNFOrig_Click
End Sub

Private Sub LabelFilialFornNFOrig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialFornNFOrig, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialFornNFOrig_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialFornNFOrig, Source, X, Y)
End Sub

Private Sub LabelFornNFOrig_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornNFOrig, Source, X, Y)
End Sub

Private Sub LabelFornNFOrig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornNFOrig, Button, Shift, X, Y)
End Sub

Private Sub BotaoGrade_Click()
    Call objCT.BotaoGrade_Click
End Sub

Public Sub ProdutoAlmoxDist_Change()
'distribuicao

    Call objCT.ProdutoAlmoxDist_Change

End Sub

Public Sub ProdutoAlmoxDist_GotFocus()
'distribuicao

    Call objCT.ProdutoAlmoxDist_GotFocus

End Sub

Public Sub ProdutoAlmoxDist_KeyPress(KeyAscii As Integer)
'distribuicao

    Call objCT.ProdutoAlmoxDist_KeyPress(KeyAscii)

End Sub

Public Sub ProdutoAlmoxDist_Validate(Cancel As Boolean)
'distribuicao

    Call objCT.ProdutoAlmoxDist_Validate(Cancel)

End Sub

Private Sub TaxaConversao_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.TaxaConversao_Change(objCT)
End Sub

Private Sub TaxaConversao_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.TaxaConversao_Validate(objCT, Cancel)
End Sub

Private Sub Historico_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Historico_Change(objCT)
End Sub

Private Sub Historico_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Historico_Click(objCT)
End Sub

Private Sub SubConta_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.SubConta_Change(objCT)
End Sub

Private Sub SubConta_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.SubConta_Click(objCT)
End Sub

Private Sub BotaoDocContrato_Click()
     Call objCT.BotaoDocContrato_Click
End Sub

Private Sub BotaoItemContrato_Click()
     Call objCT.BotaoItemContrato_Click
End Sub

Private Sub BotaoMedicao_Click()
     Call objCT.BotaoMedicao_Click
End Sub

Private Sub Item_Change()
     Call objCT.Item_Change
End Sub

Private Sub Item_GotFocus()
     Call objCT.Item_GotFocus
End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)
     Call objCT.Item_KeyPress(KeyAscii)
End Sub

Private Sub Item_Validate(Cancel As Boolean)
     Call objCT.Item_Validate(Cancel)
End Sub

Private Sub Contrato_Change()
     Call objCT.Contrato_Change
End Sub

Private Sub Contrato_GotFocus()
     Call objCT.Contrato_GotFocus
End Sub

Private Sub Contrato_KeyPress(KeyAscii As Integer)
     Call objCT.Contrato_KeyPress(KeyAscii)
End Sub

Private Sub Contrato_Validate(Cancel As Boolean)
     Call objCT.Contrato_Validate(Cancel)
End Sub

Private Sub BotaoSerie_Click()
    Call objCT.BotaoSerie_Click
End Sub

Private Sub BotaoProjetos_Click()
    Call objCT.BotaoProjetos_Click
End Sub

Private Sub LabelProjeto_Click()
    Call objCT.LabelProjeto_Click
End Sub

Private Sub Projeto_Change()
     Call objCT.Projeto_Change
End Sub

Private Sub Projeto_GotFocus()
     Call objCT.Projeto_GotFocus
End Sub

Private Sub Projeto_Validate(Cancel As Boolean)
     Call objCT.Projeto_Validate(Cancel)
End Sub

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

Private Sub CodigodeBarras_Change()
     Call objCT.CodigodeBarras_Change
End Sub

Private Sub CodigodeBarras_GotFocus()
    Call objCT.CodigodeBarras_GotFocus
End Sub

Private Sub CodigodeBarras_KeyPress(KeyAscii As Integer)
    Call objCT.CodigodeBarras_KeyPress(KeyAscii)
End Sub

Private Sub CodigodeBarras_Validate(Cancel As Boolean)
    Call objCT.CodigodeBarras_Validate(Cancel)
End Sub

Private Sub MensagemCorpo_Change()
     Call objCT.MensagemCorpo_Change
End Sub

Private Sub MsgAutomatica_Click()
     Call objCT.MsgAutomatica_Click
End Sub

Private Sub EletronicaFed_Click()
    Call objCT.EletronicaFed_Click
End Sub


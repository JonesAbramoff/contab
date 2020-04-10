VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RastreamentoLote 
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   9090
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5880
      Index           =   1
      Left            =   285
      TabIndex        =   6
      Top             =   750
      Width           =   8565
      Begin VB.Frame Frame6 
         Caption         =   "Cliente"
         Height          =   810
         Left            =   330
         TabIndex        =   69
         Top             =   4500
         Width           =   7545
         Begin VB.ComboBox FilialCli 
            Height          =   315
            Left            =   4080
            TabIndex        =   72
            Top             =   270
            Width           =   2190
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   315
            Left            =   915
            TabIndex        =   70
            Top             =   300
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelFilialCli 
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
            Height          =   195
            Left            =   3540
            TabIndex        =   73
            Top             =   330
            Width           =   480
         End
         Begin VB.Label LabelCliente 
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
            Height          =   195
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   71
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Principais"
         Height          =   1305
         Index           =   0
         Left            =   375
         TabIndex        =   20
         Top             =   120
         Width           =   7530
         Begin VB.ComboBox FilialOP 
            Height          =   315
            Left            =   5115
            TabIndex        =   21
            Top             =   300
            Width           =   2295
         End
         Begin MSMask.MaskEdBox Lote 
            Height          =   300
            Left            =   885
            TabIndex        =   22
            Top             =   300
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   870
            TabIndex        =   23
            Top             =   810
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LoteLabel 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   375
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   330
            Width           =   450
         End
         Begin VB.Label ProdutoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
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
            Left            =   105
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Descricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2865
            TabIndex        =   25
            Top             =   810
            Width           =   4530
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "FilialOP:"
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
            Left            =   4380
            TabIndex        =   24
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Complementares"
         Height          =   2835
         Left            =   345
         TabIndex        =   8
         Top             =   1605
         Width           =   7575
         Begin VB.TextBox Localizacao 
            Height          =   300
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   67
            Top             =   2445
            Width           =   4350
         End
         Begin VB.TextBox Observacao 
            Height          =   960
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1365
            Width           =   5790
         End
         Begin MSComCtl2.UpDown UpDownValidade 
            Height          =   300
            Left            =   7110
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataValidade 
            Height          =   300
            Left            =   6000
            TabIndex        =   11
            Top             =   840
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownFabricacao 
            Height          =   300
            Left            =   2670
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   855
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFabricacao 
            Height          =   300
            Left            =   1560
            TabIndex        =   13
            Top             =   855
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEntrada 
            Height          =   300
            Left            =   2670
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntrada 
            Height          =   300
            Left            =   1560
            TabIndex        =   15
            Top             =   405
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelLocalizacao 
            AutoSize        =   -1  'True
            Caption         =   "Localização:"
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
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   68
            Top             =   2460
            Width           =   1095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data Validade:"
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
            Left            =   4650
            TabIndex        =   19
            Top             =   870
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data Fabricação:"
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
            TabIndex        =   18
            Top             =   915
            Width           =   1485
         End
         Begin VB.Label Label8 
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
            Left            =   450
            TabIndex        =   17
            Top             =   1395
            Width           =   1095
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
            Left            =   345
            TabIndex        =   16
            Top             =   435
            Width           =   1200
         End
      End
      Begin VB.CommandButton BotaoRastroMovto 
         Caption         =   "Histórico dos Movimentos dos Lotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   345
         TabIndex        =   7
         Top             =   5415
         Width           =   7515
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5925
      Index           =   2
      Left            =   285
      TabIndex        =   28
      Top             =   765
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame Frame5 
         Caption         =   "Padrão"
         Height          =   645
         Left            =   60
         TabIndex        =   64
         Top             =   -30
         Width           =   8400
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   5895
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   370
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   315
            Left            =   4830
            TabIndex        =   31
            Top             =   225
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IdAnalisePadrao 
            Height          =   315
            Left            =   1260
            TabIndex        =   30
            Top             =   225
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Id. Análise:"
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
            Left            =   270
            TabIndex        =   66
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Data Análise:"
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
            Left            =   3645
            TabIndex        =   65
            Top             =   285
            Width           =   1155
         End
      End
      Begin VB.CommandButton BotaoTrazerTeste 
         Caption         =   "Trazer Testes"
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
         Left            =   1950
         TabIndex        =   60
         Top             =   5490
         Width           =   1815
      End
      Begin VB.CommandButton BotaoImprimirLaudo 
         Caption         =   "Imprimir Laudo"
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
         Left            =   75
         TabIndex        =   59
         Top             =   5490
         Width           =   1815
      End
      Begin VB.CommandButton BotaoTestes 
         Caption         =   "Testes"
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
         Left            =   7230
         TabIndex        =   61
         Top             =   5475
         Width           =   1230
      End
      Begin VB.Frame Frame4 
         Caption         =   "Informações sobre o teste selecionado acima"
         Height          =   2430
         Left            =   60
         TabIndex        =   37
         Top             =   3015
         Width           =   8415
         Begin VB.TextBox LabelObservacao 
            BackColor       =   &H8000000F&
            Height          =   780
            Left            =   4350
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   1560
            Width           =   3825
         End
         Begin VB.TextBox LabelEspecificacao 
            BackColor       =   &H8000000F&
            Height          =   780
            Left            =   165
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   1560
            Width           =   3825
         End
         Begin VB.CheckBox NoCertificado2 
            Caption         =   "O resultado deste teste deve aparecer no certificado"
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
            Left            =   3300
            TabIndex        =   57
            Top             =   270
            Value           =   1  'Checked
            Width           =   4890
         End
         Begin VB.Frame FrameLimites 
            Caption         =   "Limites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   165
            TabIndex        =   41
            Top             =   660
            Width           =   3945
            Begin VB.Label LabelLimiteAte 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2520
               TabIndex        =   55
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label LabelLimiteDe 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   570
               TabIndex        =   54
               Top             =   255
               Width           =   1050
            End
            Begin VB.Label Label6 
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
               Height          =   285
               Left            =   2040
               TabIndex        =   43
               Top             =   270
               Width           =   375
            End
            Begin VB.Label Label5 
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
               Height          =   285
               Left            =   165
               TabIndex        =   42
               Top             =   270
               Width           =   375
            End
         End
         Begin VB.Label LabelMetodo 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   5220
            TabIndex        =   56
            Top             =   870
            Width           =   2640
         End
         Begin VB.Label LabelTeste 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   810
            TabIndex        =   47
            Top             =   330
            Width           =   2115
         End
         Begin VB.Label LabelTesteCodigo 
            AutoSize        =   -1  'True
            Caption         =   "Teste:"
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   46
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label11 
            Caption         =   "Método:"
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
            Left            =   4350
            TabIndex        =   45
            Top             =   870
            Width           =   750
         End
         Begin VB.Label Label10 
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
            Height          =   180
            Left            =   4335
            TabIndex        =   44
            Top             =   1350
            Width           =   1305
         End
         Begin VB.Label Label3 
            Caption         =   "Especificação:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   165
            TabIndex        =   40
            Top             =   1350
            Width           =   1305
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resultados dos Testes"
         Height          =   2340
         Left            =   60
         TabIndex        =   29
         Top             =   645
         Width           =   8415
         Begin VB.TextBox ResultadoObs 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   1035
            MaxLength       =   250
            TabIndex        =   58
            Text            =   "Text1"
            Top             =   1650
            Width           =   2475
         End
         Begin VB.TextBox Especificacao 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   4320
            MaxLength       =   250
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   1170
            Width           =   2475
         End
         Begin VB.TextBox ObservacaoTeste 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   4335
            MaxLength       =   250
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   1560
            Width           =   2475
         End
         Begin VB.CheckBox NoCertificado 
            Caption         =   "Check1"
            Height          =   300
            Left            =   3240
            TabIndex        =   49
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox Metodo 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   4335
            MaxLength       =   50
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   1905
            Width           =   2475
         End
         Begin MSMask.MaskEdBox RegistroAnaliseData 
            Height          =   300
            Left            =   6255
            TabIndex        =   39
            Top             =   570
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox RegistroAnaliseID 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   4965
            MaxLength       =   20
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   495
            Width           =   1245
         End
         Begin VB.CheckBox ResultadoNaoConforme 
            Caption         =   "Não Conforme"
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
            Left            =   1995
            TabIndex        =   36
            Top             =   465
            Width           =   1275
         End
         Begin VB.TextBox ResultadoValor 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   2895
            MaxLength       =   250
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   480
            Width           =   1875
         End
         Begin VB.TextBox Teste 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   270
            MaxLength       =   100
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   480
            Width           =   2500
         End
         Begin MSMask.MaskEdBox LimiteDe 
            Height          =   300
            Left            =   1410
            TabIndex        =   52
            Top             =   1200
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LimiteAte 
            Height          =   300
            Left            =   2310
            TabIndex        =   53
            Top             =   1185
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridTestes 
            Height          =   1995
            Left            =   90
            TabIndex        =   33
            Top             =   255
            Width           =   8205
            _ExtentX        =   14473
            _ExtentY        =   3519
            _Version        =   393216
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6735
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "RastreamentoLote.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RastreamentoLote.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RastreamentoLote.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "RastreamentoLote.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   6360
      Left            =   150
      TabIndex        =   5
      Top             =   375
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11218
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Controle de Qualidade"
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
Attribute VB_Name = "RastreamentoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTRastreamentoLote
Attribute objCT.VB_VarHelpID = -1

Private Sub GridTestes_Click()
    Call objCT.GridTestes_Click
End Sub

Private Sub GridTestes_EnterCell()
    Call objCT.GridTestes_EnterCell
End Sub

Private Sub GridTestes_GotFocus()
    Call objCT.GridTestes_GotFocus
End Sub

Private Sub GridTestes_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.GridTestes_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridTestes_KeyPress(KeyAscii As Integer)
    Call objCT.GridTestes_KeyPress(KeyAscii)
End Sub

Private Sub GridTestes_LeaveCell()
    Call objCT.GridTestes_LeaveCell
End Sub

Private Sub GridTestes_RowColChange()
    Call objCT.GridTestes_RowColChange
End Sub

Private Sub GridTestes_Scroll()
    Call objCT.GridTestes_Scroll
End Sub

Private Sub GridTestes_Validate(Cancel As Boolean)
    Call objCT.GridTestes_Validate(Cancel)
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTRastreamentoLote
    Set objCT.objUserControl = Me
End Sub

Function Trata_Parametros(Optional objRastroLote As ClassRastreamentoLote) As Long
     Trata_Parametros = objCT.Trata_Parametros(objRastroLote)
End Function

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoImprimirLaudo_Click()
     Call objCT.BotaoImprimirLaudo_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoRastroMovto_Click()
     Call objCT.BotaoRastroMovto_Click
End Sub

Private Sub DataValidade_Validate(Cancel As Boolean)
     Call objCT.DataValidade_Validate(Cancel)
End Sub

Private Sub DataEntrada_Validate(Cancel As Boolean)
     Call objCT.DataEntrada_Validate(Cancel)
End Sub

Private Sub DataFabricacao_Validate(Cancel As Boolean)
     Call objCT.DataFabricacao_Validate(Cancel)
End Sub

Private Sub FilialOP_Change()
     Call objCT.FilialOP_Change
End Sub

Private Sub FilialOP_Validate(Cancel As Boolean)
     Call objCT.FilialOP_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub UpDownValidade_DownClick()
     Call objCT.UpDownValidade_DownClick
End Sub

Private Sub UpDownValidade_UpClick()
     Call objCT.UpDownValidade_UpClick
End Sub

Private Sub UpDownEntrada_DownClick()
     Call objCT.UpDownEntrada_DownClick
End Sub

Private Sub UpDownEntrada_UpClick()
     Call objCT.UpDownEntrada_UpClick
End Sub

Private Sub UpDownFabricacao_DownClick()
     Call objCT.UpDownFabricacao_DownClick
End Sub

Private Sub UpDownFabricacao_UpClick()
     Call objCT.UpDownFabricacao_UpClick
End Sub

Private Sub LoteLabel_Click()
     Call objCT.LoteLabel_Click
End Sub

Private Sub ProdutoLabel_Click()
     Call objCT.ProdutoLabel_Click
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub Lote_Change()
     Call objCT.Lote_Change
End Sub

Private Sub DataEntrada_Change()
     Call objCT.DataEntrada_Change
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub DataValidade_Change()
     Call objCT.DataValidade_Change
End Sub

Private Sub DataFabricacao_Change()
     Call objCT.DataFabricacao_Change
End Sub

Private Sub Observacao_Change()
     Call objCT.Observacao_Change
End Sub

Private Sub Lote_GotFocus()
     Call objCT.Lote_GotFocus
End Sub

Private Sub DataEntrada_GotFocus()
     Call objCT.DataEntrada_GotFocus
End Sub

Private Sub DataValidade_GotFocus()
     Call objCT.DataValidade_GotFocus
End Sub

Private Sub DataFabricacao_GotFocus()
     Call objCT.DataFabricacao_GotFocus
End Sub

Private Sub LoteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LoteLabel, Source, X, Y)
End Sub
Private Sub LoteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LoteLabel, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub
Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub
Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub
Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub
Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub
Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub LimiteDe_Change()
     Call objCT.LimiteDe_Change
End Sub

Private Sub LimiteDe_Click()
     Call objCT.LimiteDe_Click
End Sub

Private Sub LimiteDe_GotFocus()
     Call objCT.LimiteDe_GotFocus
End Sub

Private Sub LimiteDe_KeyPress(KeyAscii As Integer)
     Call objCT.LimiteDe_KeyPress(KeyAscii)
End Sub

Private Sub LimiteDe_Validate(Cancel As Boolean)
     Call objCT.LimiteDe_Validate(Cancel)
End Sub

Private Sub LimiteAte_Change()
     Call objCT.LimiteAte_Change
End Sub

Private Sub LimiteAte_Click()
     Call objCT.LimiteAte_Click
End Sub

Private Sub LimiteAte_GotFocus()
     Call objCT.LimiteAte_GotFocus
End Sub

Private Sub LimiteAte_KeyPress(KeyAscii As Integer)
     Call objCT.LimiteAte_KeyPress(KeyAscii)
End Sub

Private Sub LimiteAte_Validate(Cancel As Boolean)
     Call objCT.LimiteAte_Validate(Cancel)
End Sub

Private Sub NoCertificado_Change()
     Call objCT.NoCertificado_Change
End Sub

Private Sub NoCertificado_Click()
     Call objCT.NoCertificado_Click
End Sub

Private Sub NoCertificado_GotFocus()
     Call objCT.NoCertificado_GotFocus
End Sub

Private Sub NoCertificado_KeyPress(KeyAscii As Integer)
     Call objCT.NoCertificado_KeyPress(KeyAscii)
End Sub

Private Sub NoCertificado_Validate(Cancel As Boolean)
     Call objCT.NoCertificado_Validate(Cancel)
End Sub

Private Sub Especificacao_Change()
     Call objCT.Especificacao_Change
End Sub

Private Sub Especificacao_Click()
     Call objCT.Especificacao_Click
End Sub

Private Sub Especificacao_GotFocus()
     Call objCT.Especificacao_GotFocus
End Sub

Private Sub Especificacao_KeyPress(KeyAscii As Integer)
     Call objCT.Especificacao_KeyPress(KeyAscii)
End Sub

Private Sub Especificacao_Validate(Cancel As Boolean)
     Call objCT.Especificacao_Validate(Cancel)
End Sub

Private Sub ObservacaoTeste_Change()
     Call objCT.ObservacaoTeste_Change
End Sub

Private Sub ObservacaoTeste_Click()
     Call objCT.ObservacaoTeste_Click
End Sub

Private Sub ObservacaoTeste_GotFocus()
     Call objCT.ObservacaoTeste_GotFocus
End Sub

Private Sub ObservacaoTeste_KeyPress(KeyAscii As Integer)
     Call objCT.ObservacaoTeste_KeyPress(KeyAscii)
End Sub

Private Sub ObservacaoTeste_Validate(Cancel As Boolean)
     Call objCT.ObservacaoTeste_Validate(Cancel)
End Sub

Private Sub Metodo_Change()
     Call objCT.Metodo_Change
End Sub

Private Sub Metodo_Click()
     Call objCT.Metodo_Click
End Sub

Private Sub Metodo_GotFocus()
     Call objCT.Metodo_GotFocus
End Sub

Private Sub Metodo_KeyPress(KeyAscii As Integer)
     Call objCT.Metodo_KeyPress(KeyAscii)
End Sub

Private Sub Metodo_Validate(Cancel As Boolean)
     Call objCT.Metodo_Validate(Cancel)
End Sub

Private Sub BotaoTestes_Click()
     Call objCT.BotaoTestes_Click
End Sub

Private Sub Teste_Change()
     Call objCT.Teste_Change
End Sub

Private Sub Teste_Click()
     Call objCT.Teste_Click
End Sub

Private Sub Teste_GotFocus()
     Call objCT.Teste_GotFocus
End Sub

Private Sub Teste_KeyPress(KeyAscii As Integer)
     Call objCT.Teste_KeyPress(KeyAscii)
End Sub

Private Sub Teste_Validate(Cancel As Boolean)
     Call objCT.Teste_Validate(Cancel)
End Sub

Private Sub ResultadoNaoConforme_Change()
     Call objCT.ResultadoNaoConforme_Change
End Sub

Private Sub ResultadoNaoConforme_Click()
     Call objCT.ResultadoNaoConforme_Click
End Sub

Private Sub ResultadoNaoConforme_GotFocus()
     Call objCT.ResultadoNaoConforme_GotFocus
End Sub

Private Sub ResultadoNaoConforme_KeyPress(KeyAscii As Integer)
     Call objCT.ResultadoNaoConforme_KeyPress(KeyAscii)
End Sub

Private Sub ResultadoNaoConforme_Validate(Cancel As Boolean)
     Call objCT.ResultadoNaoConforme_Validate(Cancel)
End Sub

Private Sub ResultadoValor_Change()
     Call objCT.ResultadoValor_Change
End Sub

Private Sub ResultadoValor_Click()
     Call objCT.ResultadoValor_Click
End Sub

Private Sub ResultadoValor_GotFocus()
     Call objCT.ResultadoValor_GotFocus
End Sub

Private Sub ResultadoValor_KeyPress(KeyAscii As Integer)
     Call objCT.ResultadoValor_KeyPress(KeyAscii)
End Sub

Private Sub ResultadoValor_Validate(Cancel As Boolean)
     Call objCT.ResultadoValor_Validate(Cancel)
End Sub

Private Sub RegistroAnaliseID_Change()
     Call objCT.RegistroAnaliseID_Change
End Sub

Private Sub RegistroAnaliseID_Click()
     Call objCT.RegistroAnaliseID_Click
End Sub

Private Sub RegistroAnaliseID_GotFocus()
     Call objCT.RegistroAnaliseID_GotFocus
End Sub

Private Sub RegistroAnaliseID_KeyPress(KeyAscii As Integer)
     Call objCT.RegistroAnaliseID_KeyPress(KeyAscii)
End Sub

Private Sub RegistroAnaliseID_Validate(Cancel As Boolean)
     Call objCT.RegistroAnaliseID_Validate(Cancel)
End Sub

Private Sub RegistroAnaliseData_Change()
     Call objCT.RegistroAnaliseData_Change
End Sub

Private Sub RegistroAnaliseData_Click()
     Call objCT.RegistroAnaliseData_Click
End Sub

Private Sub RegistroAnaliseData_GotFocus()
     Call objCT.RegistroAnaliseData_GotFocus
End Sub

Private Sub RegistroAnaliseData_KeyPress(KeyAscii As Integer)
     Call objCT.RegistroAnaliseData_KeyPress(KeyAscii)
End Sub

Private Sub RegistroAnaliseData_Validate(Cancel As Boolean)
     Call objCT.RegistroAnaliseData_Validate(Cancel)
End Sub

Private Sub ResultadoObs_Change()
     Call objCT.ResultadoObs_Change
End Sub

Private Sub ResultadoObs_Click()
     Call objCT.ResultadoObs_Click
End Sub

Private Sub ResultadoObs_GotFocus()
     Call objCT.ResultadoObs_GotFocus
End Sub

Private Sub ResultadoObs_KeyPress(KeyAscii As Integer)
     Call objCT.ResultadoObs_KeyPress(KeyAscii)
End Sub

Private Sub ResultadoObs_Validate(Cancel As Boolean)
     Call objCT.ResultadoObs_Validate(Cancel)
End Sub

Private Sub Teste_ExibeInfo(ByVal iLinha As Integer)
     Call objCT.Teste_ExibeInfo(iLinha)
End Sub

Private Sub BotaoTrazerTeste_Click()
     Call objCT.BotaoTrazerTeste_Click
End Sub

Private Sub UpDownData_DownClick()
     Call objCT.UpDownData_DownClick
End Sub

Private Sub UpDownData_UpClick()
     Call objCT.UpDownData_UpClick
End Sub

Private Sub Data_Change()
     Call objCT.Data_Change
End Sub

Private Sub IdAnalisePadrao_Change()
     Call objCT.IdAnalisePadrao_Change
End Sub

Private Sub Data_GotFocus()
     Call objCT.Data_GotFocus
End Sub

Private Sub Data_Validate(Cancel As Boolean)
     Call objCT.Data_Validate(Cancel)
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

Private Sub FilialCli_Change()
    Call objCT.Filial_Change
End Sub

Private Sub FilialCli_Validate(Cancel As Boolean)
    Call objCT.FilialCli_Validate(Cancel)
End Sub

Private Sub Cliente_Change()
    Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
    Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub LabelCliente_Click()
    Call objCT.LabelCliente_Click
End Sub

Private Sub Localizacao_Change()
    Call objCT.Localizacao_Change
End Sub

Private Sub LabelLocalizacao_Click()
    Call objCT.LabelLocalizacao_Click
End Sub

Private Sub Localizacao_Validate(Cancel As Boolean)
    Call objCT.Localizacao_Validate(Cancel)
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ConhecimentoFrete 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleMode       =   0  'User
   ScaleWidth      =   9375.636
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4650
      Index           =   4
      Left            =   135
      TabIndex        =   48
      Top             =   930
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame Frame4 
         Caption         =   "Mercadoria(s) Transportada(s)"
         Height          =   3645
         Left            =   105
         TabIndex        =   49
         Top             =   60
         Width           =   8925
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   3270
            TabIndex        =   52
            Top             =   330
            Width           =   1335
         End
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   5490
            TabIndex        =   53
            Top             =   330
            Width           =   1335
         End
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7335
            MaxLength       =   20
            TabIndex        =   54
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox ImprimeMsgICMS 
            Caption         =   "Imprimir Mensagem de não Incidência do ICMS."
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
            Left            =   480
            TabIndex        =   60
            Top             =   2160
            Width           =   5295
         End
         Begin VB.TextBox NotasFiscais 
            Height          =   300
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   58
            Top             =   1290
            Width           =   6480
         End
         Begin VB.TextBox Observacao 
            Height          =   300
            Left            =   1590
            MaxLength       =   20
            TabIndex        =   59
            Top             =   1755
            Width           =   6480
         End
         Begin VB.Frame Frame2 
            Caption         =   "Locais"
            Height          =   900
            Index           =   0
            Left            =   450
            TabIndex        =   76
            Top             =   2520
            Width           =   8025
            Begin VB.TextBox Entrega 
               Height          =   300
               Left            =   3315
               MaxLength       =   20
               TabIndex        =   62
               Top             =   360
               Width           =   1500
            End
            Begin VB.TextBox Coleta 
               Height          =   300
               Left            =   945
               MaxLength       =   20
               TabIndex        =   61
               Top             =   360
               Width           =   1440
            End
            Begin VB.TextBox CalculadoAte 
               Height          =   300
               Left            =   6420
               MaxLength       =   20
               TabIndex        =   63
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Entrega:"
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
               Index           =   10
               Left            =   2505
               TabIndex        =   78
               Top             =   413
               Width           =   735
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Coleta:"
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
               Index           =   9
               Left            =   195
               TabIndex        =   77
               Top             =   413
               Width           =   615
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Calculado Até :"
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
               Index           =   12
               Left            =   5055
               TabIndex        =   79
               Top             =   435
               Width           =   1320
            End
         End
         Begin VB.TextBox NaturezaCarga 
            Height          =   300
            Left            =   1575
            MaxLength       =   20
            TabIndex        =   55
            Top             =   825
            Width           =   1425
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1590
            TabIndex        =   51
            Top             =   360
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoMercadoria 
            Height          =   300
            Left            =   3675
            TabIndex        =   56
            Top             =   810
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorMercadoria 
            Height          =   300
            Left            =   7350
            TabIndex        =   57
            Top             =   825
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Espécie:"
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
            Left            =   2460
            TabIndex        =   239
            Top             =   405
            Width           =   750
         End
         Begin VB.Label Label30 
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
            Height          =   195
            Index           =   2
            Left            =   4845
            TabIndex        =   238
            Top             =   405
            Width           =   600
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Notas Fiscais:"
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
            Index           =   7
            Left            =   330
            TabIndex        =   74
            Top             =   1335
            Width           =   1215
         End
         Begin VB.Label Label30 
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
            Index           =   8
            Left            =   465
            TabIndex        =   75
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Nº :"
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
            Left            =   6960
            TabIndex        =   68
            Top             =   420
            Width           =   345
         End
         Begin VB.Label Label30 
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
            Height          =   195
            Index           =   0
            Left            =   510
            TabIndex        =   50
            Top             =   420
            Width           =   1050
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
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
            Left            =   3165
            TabIndex        =   70
            Top             =   885
            Width           =   495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Kg ou m"
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
            Left            =   4665
            TabIndex        =   71
            Top             =   870
            Width           =   705
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5355
            TabIndex        =   72
            Top             =   810
            Width           =   90
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Valor Mercadoria:"
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
            Left            =   5820
            TabIndex        =   73
            Top             =   870
            Width           =   1515
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Carga:"
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
            Index           =   13
            Left            =   150
            TabIndex        =   69
            Top             =   885
            Width           =   1395
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dados Veículo"
         Height          =   780
         Index           =   1
         Left            =   105
         TabIndex        =   80
         Top             =   3810
         Width           =   8955
         Begin VB.TextBox MarcaVeiculo 
            Height          =   315
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   64
            Top             =   300
            Width           =   1290
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   3335
            MaxLength       =   10
            TabIndex        =   65
            Top             =   300
            Width           =   1290
         End
         Begin VB.TextBox LocalVeiculo 
            Height          =   315
            Left            =   5545
            MaxLength       =   15
            TabIndex        =   66
            Top             =   300
            Width           =   1290
         End
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   7755
            TabIndex        =   67
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label30 
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
            Height          =   195
            Index           =   6
            Left            =   495
            TabIndex        =   81
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Placa:"
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
            Index           =   4
            Left            =   2745
            TabIndex        =   82
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Local:"
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
            Index           =   5
            Left            =   5010
            TabIndex        =   83
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
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
            Left            =   7410
            TabIndex        =   84
            Top             =   360
            Width           =   315
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   915
      Visible         =   0   'False
      Width           =   9120
      Begin VB.Frame Frame3 
         Caption         =   "Comprovantes de Serviços"
         Height          =   2565
         Left            =   120
         TabIndex        =   22
         Top             =   60
         Width           =   8940
         Begin MSMask.MaskEdBox PedagioCon 
            Height          =   225
            Left            =   2715
            TabIndex        =   222
            Top             =   1395
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox PrecoCon 
            Height          =   225
            Left            =   5715
            TabIndex        =   221
            Top             =   1065
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorContainerCon 
            Height          =   225
            Left            =   5355
            TabIndex        =   220
            Top             =   660
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorMercadoriaCon 
            Height          =   225
            Left            =   3870
            TabIndex        =   219
            Top             =   945
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox AdValorenCon 
            Height          =   225
            Left            =   4005
            TabIndex        =   218
            Top             =   585
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox QuantCon 
            Height          =   225
            Left            =   2895
            TabIndex        =   217
            Top             =   570
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DescricaoCon 
            Height          =   225
            Left            =   1575
            TabIndex        =   216
            Top             =   930
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ServicoCon 
            Height          =   225
            Left            =   1890
            TabIndex        =   215
            Top             =   585
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataCon 
            Height          =   225
            Left            =   555
            TabIndex        =   214
            Top             =   885
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
         Begin MSMask.MaskEdBox ComprovServCon 
            Height          =   225
            Left            =   495
            TabIndex        =   213
            Top             =   570
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            Mask            =   "##########"
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoComprovante 
            Caption         =   "Comprovante"
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
            Left            =   7455
            TabIndex        =   212
            Top             =   2115
            Width           =   1365
         End
         Begin MSFlexGridLib.MSFlexGrid GridComprovServ 
            Height          =   1785
            Left            =   195
            TabIndex        =   223
            Top             =   255
            Width           =   8625
            _ExtentX        =   15214
            _ExtentY        =   3149
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Composição do Frete"
         Height          =   2055
         Left            =   135
         TabIndex        =   23
         Top             =   2655
         Width           =   8940
         Begin VB.CheckBox PedagioIncluso 
            Caption         =   "Incluir Pedágio na Base Cálculo"
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
            Left            =   2145
            TabIndex        =   25
            Top             =   375
            UseMaskColor    =   -1  'True
            Width           =   3045
         End
         Begin VB.CheckBox ICMSIncluso 
            Caption         =   "ICMS Incluso"
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
            Left            =   555
            TabIndex        =   24
            Top             =   375
            UseMaskColor    =   -1  'True
            Value           =   1  'Checked
            Width           =   1560
         End
         Begin VB.Frame SSFrame6 
            Caption         =   "INSS"
            Height          =   555
            Left            =   5280
            TabIndex        =   26
            Top             =   180
            Width           =   3150
            Begin VB.CheckBox INSSRetido 
               Caption         =   "Retido"
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
               Left            =   2100
               TabIndex        =   29
               Top             =   225
               Width           =   900
            End
            Begin MSMask.MaskEdBox ValorINSS 
               Height          =   300
               Left            =   705
               TabIndex        =   28
               Top             =   195
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
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
               Height          =   210
               Index           =   88
               Left            =   150
               TabIndex        =   27
               Top             =   240
               Width           =   510
            End
         End
         Begin MSMask.MaskEdBox FretePeso 
            Height          =   300
            Left            =   555
            TabIndex        =   34
            Top             =   1065
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox SEC 
            Height          =   300
            Left            =   3840
            TabIndex        =   36
            Top             =   1065
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Despacho 
            Height          =   300
            Left            =   5520
            TabIndex        =   37
            Top             =   1065
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Pedagio 
            Height          =   300
            Left            =   7140
            TabIndex        =   38
            Top             =   1065
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Aliquota 
            Height          =   300
            Left            =   2175
            TabIndex        =   44
            Top             =   1605
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorICMS 
            Height          =   300
            Left            =   3840
            TabIndex        =   45
            Top             =   1605
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FreteValor 
            Height          =   300
            Left            =   2175
            TabIndex        =   35
            Top             =   1065
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox OutrosValores 
            Height          =   300
            Left            =   555
            TabIndex        =   43
            Top             =   1605
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BaseCalculo 
            Height          =   300
            Left            =   5520
            TabIndex        =   46
            Top             =   1605
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
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
            Height          =   195
            Index           =   85
            Left            =   4230
            TabIndex        =   225
            Top             =   1395
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sec/Cat"
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
            Index           =   80
            Left            =   4102
            TabIndex        =   224
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Frete Peso"
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
            Index           =   78
            Left            =   735
            TabIndex        =   30
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Despacho"
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
            Index           =   81
            Left            =   5722
            TabIndex        =   32
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedágio"
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
            Index           =   82
            Left            =   7418
            TabIndex        =   33
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota"
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
            Index           =   84
            Left            =   2445
            TabIndex        =   40
            Top             =   1395
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Base Cálculo"
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
            Index           =   86
            Left            =   5595
            TabIndex        =   41
            Top             =   1395
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Frete Valor"
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
            Index           =   79
            Left            =   2340
            TabIndex        =   31
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Outros Valores"
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
            Index           =   83
            Left            =   570
            TabIndex        =   39
            Top             =   1395
            Width           =   1260
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   87
            Left            =   7545
            TabIndex        =   42
            Top             =   1395
            Width           =   450
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7140
            TabIndex        =   47
            Top             =   1605
            Width           =   1290
         End
      End
      Begin VB.Frame FrameOculto 
         Caption         =   "Frame3"
         Height          =   1065
         Left            =   3975
         TabIndex        =   197
         Top             =   3345
         Visible         =   0   'False
         Width           =   1575
         Begin VB.TextBox ValorFrete 
            Height          =   315
            Left            =   165
            TabIndex        =   201
            Top             =   255
            Width           =   435
         End
         Begin VB.TextBox IPIValor1 
            Height          =   315
            Left            =   165
            TabIndex        =   200
            Top             =   570
            Width           =   435
         End
         Begin VB.TextBox ValorDespesas 
            Height          =   315
            Left            =   600
            TabIndex        =   199
            Top             =   570
            Width           =   435
         End
         Begin VB.TextBox ValorSeguro 
            Height          =   315
            Left            =   600
            TabIndex        =   198
            Top             =   255
            Width           =   435
         End
         Begin MSMask.MaskEdBox ISSValor 
            Height          =   300
            Left            =   1050
            TabIndex        =   209
            Top             =   585
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1050
            TabIndex        =   208
            Top             =   255
            Width           =   405
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4680
      Index           =   6
      Left            =   105
      TabIndex        =   154
      Top             =   915
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   5100
         TabIndex        =   235
         Tag             =   "1"
         Top             =   2745
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
         Left            =   7800
         TabIndex        =   162
         Top             =   90
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6375
         Style           =   2  'Dropdown List
         TabIndex        =   165
         Top             =   945
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
         Left            =   6345
         TabIndex        =   161
         Top             =   90
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
         Left            =   6345
         TabIndex        =   163
         Top             =   405
         Width           =   2700
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4800
         TabIndex        =   180
         Top             =   1680
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
         Height          =   2790
         Left            =   6360
         TabIndex        =   192
         Top             =   1605
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   186
         Top             =   3465
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   190
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   188
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
            Height          =   195
            Left            =   1125
            TabIndex        =   187
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
            Height          =   195
            Left            =   240
            TabIndex        =   189
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   181
         Top             =   2175
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   182
         Top             =   2565
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
         Left            =   3495
         TabIndex        =   174
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   176
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
         TabIndex        =   179
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
         TabIndex        =   178
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
         Left            =   1560
         TabIndex        =   177
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
         TabIndex        =   168
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
         TabIndex        =   160
         Top             =   510
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
         TabIndex        =   159
         Top             =   120
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
         Left            =   3825
         TabIndex        =   158
         Top             =   105
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6360
         TabIndex        =   193
         Top             =   1605
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
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   0
         TabIndex        =   175
         Top             =   1200
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
         TabIndex        =   194
         Top             =   1605
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
         Left            =   6405
         TabIndex        =   164
         Top             =   720
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
         Left            =   30
         TabIndex        =   155
         Top             =   150
         Width           =   720
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
         TabIndex        =   166
         Top             =   150
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
         Height          =   195
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   157
         Top             =   150
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
         TabIndex        =   167
         Top             =   540
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   184
         Top             =   3135
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   185
         Top             =   3135
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
         Height          =   225
         Left            =   1800
         TabIndex        =   183
         Top             =   3150
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
         Left            =   6375
         TabIndex        =   191
         Top             =   1350
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
         Left            =   6375
         TabIndex        =   195
         Top             =   1350
         Width           =   2340
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
         Left            =   6375
         TabIndex        =   196
         Top             =   1350
         Visible         =   0   'False
         Width           =   1005
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
         TabIndex        =   173
         Top             =   930
         Width           =   1140
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
         TabIndex        =   169
         Top             =   570
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   170
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5040
         TabIndex        =   172
         Top             =   555
         Width           =   1185
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
         TabIndex        =   171
         Top             =   585
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   156
         Top             =   105
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4710
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   915
      Width           =   9195
      Begin VB.Frame Frame8 
         Caption         =   "Datas"
         Height          =   1005
         Left            =   90
         TabIndex        =   17
         Top             =   3465
         Width           =   8940
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   2025
            TabIndex        =   19
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   3105
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   450
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data de Emissão:"
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
            Index           =   4
            Left            =   450
            TabIndex        =   18
            Top             =   480
            Width           =   1500
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dados do Cliente"
         Height          =   1095
         Left            =   90
         TabIndex        =   12
         Top             =   2250
         Width           =   8940
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   6015
            TabIndex        =   16
            Top             =   495
            Width           =   1845
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1980
            TabIndex        =   14
            Top             =   450
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Index           =   1
            Left            =   5520
            TabIndex        =   15
            Top             =   555
            Width           =   450
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
            Left            =   1305
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   13
            Top             =   510
            Width           =   660
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Identificação"
         Height          =   2115
         Left            =   90
         TabIndex        =   2
         Top             =   30
         Width           =   8940
         Begin VB.CommandButton BotaoLimparNF 
            Height          =   300
            Left            =   6855
            Picture         =   "ConhecimentoFreteGR.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Limpar o Número"
            Top             =   1395
            Width           =   345
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   1935
            TabIndex        =   8
            Top             =   1395
            Width           =   765
         End
         Begin MSMask.MaskEdBox NatOpInterna 
            Height          =   300
            Left            =   1932
            TabIndex        =   4
            Top             =   648
            Width           =   624
            _ExtentX        =   1085
            _ExtentY        =   529
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
         Begin VB.Label NFiscalLabel 
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
            Height          =   255
            Left            =   5385
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   9
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label SerieLabel 
            AutoSize        =   -1  'True
            Caption         =   "Série:"
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
            Left            =   1365
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   7
            Top             =   1455
            Width           =   510
         End
         Begin VB.Label NFiscal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6120
            TabIndex        =   10
            Top             =   1395
            Width           =   735
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
            Height          =   195
            Index           =   5
            Left            =   5340
            TabIndex        =   5
            Top             =   705
            Width           =   615
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6015
            TabIndex        =   6
            Top             =   645
            Width           =   1080
         End
         Begin VB.Label LblNatOpInterna 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Operação:"
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
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   3
            Top             =   705
            Width           =   1725
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4725
      Index           =   5
      Left            =   120
      TabIndex        =   142
      Top             =   900
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame SSFrame4 
         Caption         =   "Totais - Comissões"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   227
         Top             =   3780
         Width           =   6855
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   20
            Left            =   4920
            TabIndex        =   233
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Percentual:"
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
            Index           =   19
            Left            =   2760
            TabIndex        =   232
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label TotalPercentualComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3840
            TabIndex        =   231
            Top             =   360
            Width           =   855
         End
         Begin VB.Label TotalValorComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5520
            TabIndex        =   230
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Base:"
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
            Index           =   18
            Left            =   120
            TabIndex        =   229
            Top             =   360
            Width           =   990
         End
         Begin VB.Label TotalValorBase 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1200
            TabIndex        =   228
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CheckBox ComissaoAutomatica 
         Caption         =   "Calcula comissão automaticamente"
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
         Left            =   285
         TabIndex        =   143
         Top             =   120
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   4380
         Index           =   0
         Left            =   105
         TabIndex        =   144
         Top             =   360
         Width           =   9045
         Begin VB.CommandButton BotaoVendedores 
            Caption         =   "Vendedores"
            Enabled         =   0   'False
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
            Left            =   7185
            Picture         =   "ConhecimentoFreteGR.ctx":0532
            Style           =   1  'Graphical
            TabIndex        =   234
            Top             =   3615
            Width           =   1650
         End
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   225
            Left            =   3690
            TabIndex        =   149
            Top             =   360
            Width           =   1080
            _ExtentX        =   1905
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
            Left            =   2490
            TabIndex        =   148
            Top             =   375
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   225
            Left            =   1800
            TabIndex        =   147
            Top             =   390
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   225
            Left            =   450
            TabIndex        =   146
            Top             =   375
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   225
            Left            =   5505
            TabIndex        =   151
            Top             =   390
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   225
            Left            =   4830
            TabIndex        =   150
            Top             =   375
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorBaixa 
            Height          =   225
            Left            =   7365
            TabIndex        =   153
            Top             =   345
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
         Begin MSMask.MaskEdBox PercentualBaixa 
            Height          =   225
            Left            =   6675
            TabIndex        =   152
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   1845
            Left            =   90
            TabIndex        =   145
            Top             =   420
            Width           =   8430
            _ExtentX        =   14870
            _ExtentY        =   3254
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   3
      Left            =   90
      TabIndex        =   85
      Top             =   930
      Visible         =   0   'False
      Width           =   9120
      Begin VB.Frame FrameEndereco 
         Caption         =   "Dados Remetente"
         Height          =   2280
         Index           =   0
         Left            =   225
         TabIndex        =   86
         Top             =   135
         Width           =   8370
         Begin VB.CommandButton BotaoLimpaRemetente 
            Height          =   330
            Left            =   7875
            Picture         =   "ConhecimentoFreteGR.ctx":0ADC
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Limpar"
            Top             =   285
            Width           =   390
         End
         Begin VB.ComboBox UFRemetente 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            TabIndex        =   94
            Top             =   1320
            Width           =   630
         End
         Begin VB.TextBox EnderecoRemetente 
            Height          =   315
            Left            =   1170
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   90
            Top             =   825
            Width           =   6345
         End
         Begin MSMask.MaskEdBox CidadeRemetente 
            Height          =   315
            Left            =   1170
            TabIndex        =   92
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEPRemetente 
            Height          =   315
            Left            =   6540
            TabIndex        =   96
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CGCRemetente 
            Height          =   315
            Left            =   1170
            TabIndex        =   98
            Top             =   1800
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Remetente 
            Height          =   315
            Left            =   1170
            TabIndex        =   88
            Top             =   330
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscEstRemetente 
            Height          =   315
            Left            =   5190
            TabIndex        =   100
            Top             =   1800
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            Left            =   195
            TabIndex        =   89
            Top             =   885
            Width           =   915
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   420
            TabIndex        =   91
            Top             =   1380
            Width           =   690
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Left            =   3450
            TabIndex        =   93
            Top             =   1380
            Width           =   675
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "CGC/CPF:"
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
            Left            =   240
            TabIndex        =   97
            Top             =   1860
            Width           =   885
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Insc. Estadual:"
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
            Left            =   3870
            TabIndex        =   99
            Top             =   1860
            Width           =   1290
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Left            =   6060
            TabIndex        =   95
            Top             =   1380
            Width           =   465
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Remetente:"
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
            Index           =   0
            Left            =   135
            TabIndex        =   87
            Top             =   405
            Width           =   990
         End
      End
      Begin VB.Frame FrameEndereco 
         Caption         =   "Dados Destinatário"
         Height          =   2175
         Index           =   1
         Left            =   240
         TabIndex        =   102
         Top             =   2490
         Width           =   8370
         Begin VB.CommandButton BotaoLimpaDestinatario 
            Height          =   330
            Left            =   7830
            Picture         =   "ConhecimentoFreteGR.ctx":100E
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Limpar"
            Top             =   330
            Width           =   390
         End
         Begin VB.TextBox EnderecoDestinatario 
            Height          =   315
            Left            =   1230
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   106
            Top             =   810
            Width           =   6345
         End
         Begin VB.ComboBox UFDestinatario 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4275
            TabIndex        =   110
            Top             =   1290
            Width           =   630
         End
         Begin MSMask.MaskEdBox CidadeDestinatario 
            Height          =   315
            Left            =   1245
            TabIndex        =   108
            Top             =   1290
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEPDestinatario 
            Height          =   315
            Left            =   6585
            TabIndex        =   112
            Top             =   1290
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CGCDestinatario 
            Height          =   315
            Left            =   1245
            TabIndex        =   114
            Top             =   1755
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Destinatario 
            Height          =   315
            Left            =   1230
            TabIndex        =   104
            Top             =   330
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscEstDestinatario 
            Height          =   315
            Left            =   5235
            TabIndex        =   116
            Top             =   1755
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Destinatário:"
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
            Index           =   3
            Left            =   75
            TabIndex        =   103
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Left            =   6120
            TabIndex        =   111
            Top             =   1350
            Width           =   465
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Insc. Estadual:"
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
            Left            =   3915
            TabIndex        =   115
            Top             =   1815
            Width           =   1290
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "CGC/CPF:"
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
            Left            =   285
            TabIndex        =   113
            Top             =   1815
            Width           =   885
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Left            =   3510
            TabIndex        =   109
            Top             =   1350
            Width           =   675
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   480
            TabIndex        =   107
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            Left            =   255
            TabIndex        =   105
            Top             =   900
            Width           =   915
         End
      End
   End
   Begin VB.ComboBox DiretoIndireto 
      Height          =   315
      ItemData        =   "ConhecimentoFreteGR.ctx":1540
      Left            =   6540
      List            =   "ConhecimentoFreteGR.ctx":154A
      Style           =   2  'Dropdown List
      TabIndex        =   226
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "x"
      Height          =   4710
      Index           =   8
      Left            =   135
      TabIndex        =   118
      Top             =   990
      Visible         =   0   'False
      Width           =   9150
      Begin VB.ComboBox CondicaoPagamento 
         Height          =   315
         Left            =   1440
         TabIndex        =   120
         Top             =   165
         Width           =   1815
      End
      Begin VB.CheckBox CobrancaAutomatica 
         Caption         =   "Calcula cobrança automaticamente"
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
         Left            =   3870
         TabIndex        =   121
         Top             =   225
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   3765
         Left            =   120
         TabIndex        =   122
         Top             =   705
         Width           =   8910
         Begin VB.ComboBox CarteiraCobrador 
            Height          =   315
            Left            =   5100
            Style           =   2  'Dropdown List
            TabIndex        =   211
            Top             =   2400
            Width           =   1815
         End
         Begin VB.ComboBox Cobrador 
            Height          =   315
            Left            =   3060
            Style           =   2  'Dropdown List
            TabIndex        =   210
            Top             =   2400
            Width           =   1875
         End
         Begin VB.CommandButton BotaoDataReferenciaDown 
            Height          =   150
            Left            =   3315
            Picture         =   "ConhecimentoFreteGR.ctx":1560
            Style           =   1  'Graphical
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   450
            Width           =   240
         End
         Begin VB.CommandButton BotaoDataReferenciaUp 
            Height          =   150
            Left            =   3315
            Picture         =   "ConhecimentoFreteGR.ctx":15BA
            Style           =   1  'Graphical
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
         End
         Begin VB.ComboBox Desconto1Codigo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "ConhecimentoFreteGR.ctx":1614
            Left            =   3135
            List            =   "ConhecimentoFreteGR.ctx":1616
            TabIndex        =   130
            Top             =   1140
            Width           =   1860
         End
         Begin VB.ComboBox Desconto2Codigo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3135
            TabIndex        =   134
            Top             =   1500
            Width           =   1860
         End
         Begin VB.ComboBox Desconto3Codigo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   138
            Top             =   1935
            Width           =   1860
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7425
            TabIndex        =   133
            Top             =   1140
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Valor 
            Height          =   225
            Left            =   6090
            TabIndex        =   140
            Top             =   1905
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Ate 
            Height          =   225
            Left            =   4950
            TabIndex        =   139
            Top             =   1890
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
         Begin MSMask.MaskEdBox Desconto2Valor 
            Height          =   225
            Left            =   6135
            TabIndex        =   136
            Top             =   1485
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Ate 
            Height          =   225
            Left            =   4980
            TabIndex        =   135
            Top             =   1470
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
         Begin MSMask.MaskEdBox Desconto1Valor 
            Height          =   225
            Left            =   6135
            TabIndex        =   132
            Top             =   1155
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto1Ate 
            Height          =   225
            Left            =   4935
            TabIndex        =   131
            Top             =   1140
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
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   645
            TabIndex        =   128
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
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   1815
            TabIndex        =   129
            Top             =   1155
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Percentual 
            Height          =   225
            Left            =   7425
            TabIndex        =   137
            Top             =   1485
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Percentual 
            Height          =   225
            Left            =   7365
            TabIndex        =   141
            Top             =   1920
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2880
            Left            =   270
            TabIndex        =   127
            Top             =   675
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   5080
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox DataReferencia 
            Height          =   300
            Left            =   2220
            TabIndex        =   124
            Top             =   300
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
            Caption         =   "Data de Referência:"
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
            Left            =   420
            TabIndex        =   123
            Top             =   345
            Width           =   1740
         End
      End
      Begin VB.Label CondPagtoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cond Pagto:"
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
         TabIndex        =   119
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   525
      Left            =   6525
      ScaleHeight     =   465
      ScaleWidth      =   2745
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   45
      Width           =   2805
      Begin VB.CommandButton BotaoConsultaTitRec 
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
         Left            =   75
         Picture         =   "ConhecimentoFreteGR.ctx":1618
         Style           =   1  'Graphical
         TabIndex        =   203
         ToolTipText     =   "Consulta de Título a Receber"
         Top             =   75
         Width           =   765
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   915
         Picture         =   "ConhecimentoFreteGR.ctx":1E9A
         Style           =   1  'Graphical
         TabIndex        =   204
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   1845
         Picture         =   "ConhecimentoFreteGR.ctx":1FF4
         Style           =   1  'Graphical
         TabIndex        =   206
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   2310
         Picture         =   "ConhecimentoFreteGR.ctx":2526
         Style           =   1  'Graphical
         TabIndex        =   207
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   1380
         Picture         =   "ConhecimentoFreteGR.ctx":26A4
         Style           =   1  'Graphical
         TabIndex        =   205
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   390
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5160
      Left            =   45
      TabIndex        =   0
      Top             =   570
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   9102
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comprovantes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Transporte"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complem."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label ICMSSubstBase1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   0
      TabIndex        =   237
      Top             =   15
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label ICMSSubstValor1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   236
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "ConhecimentoFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'???? Já existe na tela de Cotacao Moeda.
Const MOEDA_DOLAR = 1

'Já existe na tela de CompServico
Const TIPO_CONHECIMENTO = 1

Private Type typeCotacaoMoeda
    dValor As Double
    iMoeda As Integer
    dtData As Date
End Type

Const STRING_ITENSNF_CONTACONTABIL = 20

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFretePesoAlterado As Integer
Dim iFreteValorAlterado As Integer
Dim iSECAlterado As Integer
Dim iDespachoAlterado As Integer
Dim iPedagioAlterado As Integer
Dim iOutrosAlterado As Integer
Dim iAliquotaAlterada As Integer
Dim iICMSAlterada As Integer
Dim iBaseCalculoAlterada As Integer
Dim iClienteAlterado As Integer
Dim iValorINSSAlterado As Integer
Dim iValorContainerAlterado As Integer
Dim giDataReferenciaAlterada As Integer

'********
'Variaveis e Constatanets relacionadas ao GridComprovServ
Public objGridComprovServ As New AdmGrid

Public iGrid_ComprovServCon_col As Integer
Public iGrid_DataCon_Col As Integer
Public iGrid_ServicoCon_Col As Integer
Public iGrid_DescricaoCon_Col As Integer
Public iGrid_QuantCon_Col As Integer
Public iGrid_PrecoCon_Col As Integer
Public iGrid_AdValorenCon_Col As Integer
Public iGrid_ValorMercadoriaCon_Col As Integer
Public iGrid_ValorContainerCon_Col As Integer
Public iGrid_PedagioCon_Col As Integer

Const NUM_MAXIMO_CONHECFRETE = 240

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean

Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil

Public objGridParcelas As AdmGrid
Dim iGrid_Vencimento_col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_Cobranca_Col As Integer
Dim iGrid_CarteiraCobranca_Col As Integer
Dim iGrid_Desc1Codigo_Col As Integer
Dim iGrid_Desc1Ate_Col As Integer
Dim iGrid_Desc1Valor_Col As Integer
Dim iGrid_Desc1Percentual_Col As Integer
Dim iGrid_Desc2Codigo_Col As Integer
Dim iGrid_Desc2Ate_Col As Integer
Dim iGrid_Desc2Valor_Col As Integer
Dim iGrid_Desc2Percentual_Col As Integer
Dim iGrid_Desc3Codigo_Col As Integer
Dim iGrid_Desc3Ate_Col As Integer
Dim iGrid_Desc3Valor_Col As Integer
Dim iGrid_Desc3Percentual_Col As Integer

Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaInterna As AdmEvento
Attribute objEventoNaturezaInterna.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Public WithEvents objEventoComprovServ As AdmEvento
Attribute objEventoComprovServ.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_DADOSTRANSPORTE = 2
Private Const TAB_Complemento = 3
Private Const TAB_COBRANCA = 4
Private Const TAB_COMISSOES = 5
Private Const TAB_Contabilizacao = 6

Private Const NATUREZAOP_PADRAO_CONHECIMENTOFRETE = "562"

'************** TRATAMENTO COMISSÕES ******************
'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes

Public objGridComissoes As AdmGrid
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
'******************************************************

'Mnemônicos da Contabilidade
Const FRETE_PESO As String = "Frete_Peso"
Const FRETE_VALOR As String = "Frete_Valor"
Const VALOR_TOTAL As String = "Valor_Total"
Const SECCAT As String = "Sec_Cat"
Const DESPACHO1 As String = "Despacho"
Const PEDAGIO1 As String = "Pedagio"
Const OUTROS_VALORES As String = "Outros_Valores"
Const ICMS As String = "ICMS"
Const ICMS_INCLUSO As String = "ICMS_Incluso"
Const BASE_CALCULO As String = "Base_Calculo"
Const CLIENTE_CODIGO As String = "Cliente_Codigo"
Const CLIENTE_RAZAOSOCIAL  As String = "Cliente_Razao_Social"
Const CTA_VENDAS As String = "CtaVendas"

Public iAlterado As Integer
Dim iFrameAtual As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
   
    NatOpInterna.Text = NaturezaOp_Conv34(NATUREZAOP_PADRAO_CONHECIMENTOFRETE, gdtDataAtual)
    
    'Inicializa as Variáveis de browse
    Set objEventoSerie = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoCondPagto = New AdmEvento
    Set objGridParcelas = New AdmGrid
    Set objEventoComprovServ = New AdmEvento
    giDataReferenciaAlterada = 0
    
'************ TRATAMENTO COMISSOES *****************

        Set objGridComissoes = New AdmGrid
        Set objEventoVendedor = New AdmEvento

        Set objTabComissoes.objTela = Me

        'Inicializa o Grid de Comissões
        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError 42129

'***************************************************
'    'Carrega a Combo TipoNFiscal
'    lErro = Carrega_TipoNFiscal()
'    If lErro <> SUCESSO Then gError 84766

    'Carrega as Séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 84847
    
'    'Carrega na combo de Banco Cobradores
'    lErro = Carrega_Cobradores()
'    If lErro <> SUCESSO Then gError 95207

    'Carrega os Estados
    lErro = Carrega_PlacaUF()
    If lErro <> SUCESSO Then gError 62872
    
    'Incluído por Luiz Nogueira em 21/08/03
    'Carrega a combo VolumeEspecie
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
    If lErro <> SUCESSO Then gError 99125

    'Incluído por Luiz Nogueira em 21/08/03
    'Carrega a combo VolumeMarca
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    If lErro <> SUCESSO Then gError 99125
    
    'Inicializa Grid de Comprovante
    lErro = Inicializa_Grid_ComprovServ(objGridComprovServ)
    If lErro <> SUCESSO Then gError 99125

    'Inicializa a máscara do Produto do grid de Comprovante
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ServicoCon)
    If lErro <> SUCESSO Then gError 99126

'    'Inicializa o Grid de Parcelas
'    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
'    If lErro <> SUCESSO Then gError 62873
'
'    'Carrega na combo as Condições de Pagamento
'    lErro = Carrega_CondicaoPagamento()
'    If lErro <> SUCESSO Then gError 62874
'
'    lErro = Carrega_TipoDesconto()
'    If lErro <> SUCESSO Then gError 62875

    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO)
    If lErro <> SUCESSO Then gError 62876
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 42129, 62872, 84847, 62873, 62874, 62875, 62876, 84755, 84766, 95207, 99125, 99126

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub Aliquota_Change()
    iAlterado = REGISTRO_ALTERADO
    iAliquotaAlterada = REGISTRO_ALTERADO
End Sub

Private Sub Aliquota_GotFocus()
    iAliquotaAlterada = 0
End Sub

Private Sub Aliquota_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Aliquota_Validate

    If iAliquotaAlterada = 0 Then Exit Sub

    If Len(Aliquota.Text) > 0 Then
        'Testa o valor
        lErro = Porcentagem_Critica2(Aliquota.Text)
        If lErro <> SUCESSO Then gError 62877
    Else
        ValorICMS.Text = ""
    End If

    Call ValorTotal_Calcula

    Exit Sub

Erro_Aliquota_Validate:

    Cancel = True

    Select Case gErr

        Case 62877

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BaseCalculo_Change()

    iAlterado = REGISTRO_ALTERADO
    iBaseCalculoAlterada = REGISTRO_ALTERADO

End Sub

Private Sub BaseCalculo_GotFocus()
    iBaseCalculoAlterada = 0
End Sub

Private Sub BaseCalculo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaseCalculo_Validate

    If iBaseCalculoAlterada = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(BaseCalculo.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(BaseCalculo.Text)
        If lErro <> SUCESSO Then gError 62878

        'Põe o valor formatado na tela
        BaseCalculo.Text = Format(BaseCalculo.Text, "Fixed")

        lErro = BaseCalculo_Calula_ValorTotal
        If lErro <> SUCESSO Then gError 62879

    End If

    Exit Sub

Erro_BaseCalculo_Validate:

    Cancel = True

    Select Case gErr

        Case 62878, 62879

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'Cyntia
Sub BotaoComprovante_Click()
'Chama o browser de Comprovante
'Este Browser traz os dados do comprovante que não estejam associados
'a uma nota fiscal, que estejam abertos ou faturáveis e que a origem e o
'destino sejam diferentes

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprovante As New ClassCompServ
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoComprovante_Click
    
    If Len(Trim(Cliente.Text)) = 0 Then gError 99135
    
    objCliente.sNomeReduzido = Cliente.Text
    
    lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
    If lErro <> SUCESSO Then gError 99136
    
    'Verifica se o browser está sendo chamado do controle(F3)ou pelo grid
    If Not (Me.ActiveControl Is ComprovServCon) Then

        'Verifica se tem alguma linha selecionada no Grid
        If GridComprovServ.Row = 0 Then gError 99137

        objComprovante.lCodigo = StrParaLong(ComprovServCon.Text)

    End If
    
    colSelecao.Add objCliente.lCodigo
    
    'Chama a tela de browse de Comprovante
    Call Chama_Tela("ComprovanteFreteLista", colSelecao, objComprovante, objEventoComprovServ)

    Exit Sub

Erro_BotaoComprovante_Click:

    Select Case gErr
        
        Case 99135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 99136
            
        Case 99137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub GridComprovServ_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComprovServ, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComprovServ, iAlterado)
    End If

End Sub

Public Sub GridComprovServ_EnterCell()

    Call Grid_Entrada_Celula(objGridComprovServ, iAlterado)

End Sub

Public Sub GridComprovServ_GotFocus()

    Call Grid_Recebe_Foco(objGridComprovServ)

End Sub

Public Sub GridComprovServ_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComprovServ, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComprovServ, iAlterado)
    End If

End Sub

Public Sub GridComprovServ_LeaveCell()

    Call Saida_Celula(objGridComprovServ)

End Sub

Public Sub GridComprovServ_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComprovServ)

End Sub

Public Sub GridComprovServ_RowColChange()

    Call Grid_RowColChange(objGridComprovServ)

End Sub

Public Sub GridComprovServ_Scroll()
       
    Call Grid_Scroll(objGridComprovServ)

End Sub

'cyntia
Public Sub GridComprovServ_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim lErro As Long
Dim objComprovante As New ClassCompServ

On Error GoTo Erro_GridComprovServ_KeyDown
        
    Call Grid_Trata_Tecla1(KeyCode, objGridComprovServ)
    
    If KeyCode = vbKeyDelete Then
    
        lErro = Rotina_Comprovantes_Tela(objComprovante)
        If lErro <> SUCESSO Then gError 99066
    
    End If
    
    Exit Sub
    
Erro_GridComprovServ_KeyDown:

    Select Case gErr

        Case 99066

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
End Sub

'Cyntia
Private Sub objEventoComprovServ_evSelecao(obj1 As Object)

Dim objComprovante As ClassCompServ
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_objEventoComprovServ_evSelecao

    Set objComprovante = obj1
    
    'Verifica se alguma linha está selecionada
    If GridComprovServ.Row < 1 Then Exit Sub
    
    For iIndice = 1 To objGridComprovServ.iLinhasExistentes
        If iIndice <> GridComprovServ.Row Then
            If objComprovante.lCodigo = GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col) Then gError 99138
        End If
    Next

    ComprovServCon.PromptInclude = False
    ComprovServCon.Text = objComprovante.lCodigo
    ComprovServCon.PromptInclude = True

    If Not (Me.ActiveControl Is ComprovServCon) Then

        'Atualiza o total de linhas existentes no grid
        If GridComprovServ.Row - GridComprovServ.FixedRows = objGridComprovServ.iLinhasExistentes Then
            objGridComprovServ.iLinhasExistentes = objGridComprovServ.iLinhasExistentes + 1
        End If

        GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_ComprovServCon_col) = ComprovServCon.Text
    
        'Faz o Tratamento do Comprovante
        lErro = Rotina_Comprovantes_Tela(objComprovante)
        If lErro <> SUCESSO Then gError 99139
    
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoComprovServ_evSelecao:

    Select Case gErr
       
        Case 99138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPROVANTE_JA_EXISTENTE", gErr, objComprovante.lCodigo, iIndice)
        
        Case 99139

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim lTransacao As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoExcluir_Click
    
    
    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Serie.Text)) = 0 Then gError 89143
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 89144

'================== DADOS IDENTIFICACAO =====================
    If Len(Trim(Cliente.ClipText)) > 0 Then
        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 62971

        'Não encontrou p Cliente --> erro
        If lErro = 12348 Then gError 62972

        objNFiscal.lCliente = objCliente.lCodigo

    End If
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Serie.Text
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = TIPODOCINFO_CONHECIMENTOFRETE

    'Verifica se a existe nota fiscal está cadastrada
    lErro = CF("NFiscal_Le_1", objNFiscal)
    If lErro <> SUCESSO And lErro <> 83971 Then gError 89145

    'se a nota não está cadastrada ==> erro
    If lErro = 83971 Then gError 89146

    'pede confirmacao
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NFISCAL", objNFiscal.lNumNotaFiscal)
    If vbMsg = vbYes Then

        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 92858

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 89147
        
        'Limpa a Tela
        lErro = Limpa_Tela_NFiscal1()
        If lErro <> SUCESSO Then gError 89148

    End If

    GL_objMDIForm.MousePointer = vbDefault

    'Confirma Transação
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 89143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 89144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 89145, 89147, 89148, 92858

        Case 89146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoLimpaDestinatario_Click()

    Call Limpa_Destinatario

End Sub

Private Sub BotaoLimpaRemetente_Click()

    Call Limpa_Remetente

End Sub

Private Sub CalculadoAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Coleta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Command1_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal


objNFiscal.lNumIntDoc = 12

lErro = Trata_Parametros(objNFiscal)
If lErro <> SUCESSO Then gError 1

End Sub
Public Sub ComprovServCon_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComprovServCon_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComprovServ)

    Call MaskEdBox_TrataGotFocus(ComprovServCon, iAlterado)

End Sub

Public Sub ComprovServCon_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComprovServ)

End Sub

Public Sub ComprovServCon_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComprovServ.objControle = ComprovServCon
    lErro = Grid_Campo_Libera_Foco(objGridComprovServ)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Despacho_Change()
    iAlterado = REGISTRO_ALTERADO
    iDespachoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Despacho_GotFocus()
    iDespachoAlterado = 0
End Sub

Private Sub Despacho_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Despacho_Validate

    If iDespachoAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(Despacho.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(Despacho.Text)
        If lErro <> SUCESSO Then gError 62880

        'Põe o valor formatado na tela
        Despacho.Text = Format(Despacho.Text, "Fixed")

    End If

    Call ValorTotal_Calcula

    Exit Sub

Erro_Despacho_Validate:

    Cancel = True

    Select Case gErr

        Case 62880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Entrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub





Private Sub FretePeso_Change()
    iAlterado = REGISTRO_ALTERADO
    iFretePesoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FretePeso_GotFocus()
    iFretePesoAlterado = 0
End Sub

Private Sub FretePeso_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FretePeso_Validate

    If iFretePesoAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(FretePeso.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(FretePeso.Text)
        If lErro <> SUCESSO Then gError 62881

        'Põe o valor formatado na tela
        FretePeso.Text = Format(FretePeso.Text, "Fixed")

    End If

    Call ValorTotal_Calcula

    Exit Sub

Erro_FretePeso_Validate:

    Cancel = True

    Select Case gErr

        Case 62881

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub




Private Sub FreteValor_Change()
    iAlterado = REGISTRO_ALTERADO
    iFreteValorAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FreteValor_GotFocus()
    iFreteValorAlterado = 0
End Sub

Private Sub FreteValor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FreteValor_Validate

    If iFreteValorAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(FreteValor.ClipText)) > 0 Then

        'Critica se é valor Positivo
        lErro = Valor_Positivo_Critica_Double(FreteValor.Text)
        If lErro <> SUCESSO Then gError 62882

        'Põe o valor formatado na tela
        FreteValor.Text = Format(FreteValor.Text, "Fixed")
    End If

    Call ValorTotal_Calcula

    Exit Sub

Erro_FreteValor_Validate:

    Cancel = True

    Select Case gErr

        Case 62882

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ICMSIncluso_Click()

    iAlterado = REGISTRO_ALTERADO
    Call ValorTotal_Calcula

End Sub

Private Sub LblNatOpInterna_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection

    'Se NaturezaOP estiver preenchida coloca no Obj
    objNaturezaOp.sCodigo = NatOpInterna.Text

    colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
    colSelecao.Add NATUREZA_SAIDA_COD_FINAL

    'Chama a Tela de browse de NaturezaOp p/naturezas de entrada
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNatureza)

    Exit Sub

End Sub

Private Sub LimpaDestinatario_Click()

    Call Limpa_Destinatario

End Sub

Private Sub LimpaRemetente_Click()

    Call Limpa_Remetente

End Sub

Private Sub LocalVeiculo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Sub MarcaVeiculo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub NatOpInterna_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NatOpInterna_Validate(Cancel As Boolean)

Dim lErro As Long, objNaturezaOp As New ClassNaturezaOp

On Error GoTo Erro_NatOpInterna_Validate

    If Len(Trim(NatOpInterna.Text)) > 0 Then

        objNaturezaOp.sCodigo = NatOpInterna

        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then gError 62883

        'se nao encontrou a natureza
        If lErro <> SUCESSO Then gError 62884

        If CInt(objNaturezaOp.sCodigo) <= 500 Then gError 62885

    End If

    Exit Sub

Erro_NatOpInterna_Validate:

    Cancel = True

    Select Case gErr

        Case 62883

        Case 62884
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, NatOpInterna.Text)

        Case 62885
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr, NatOpInterna.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub NaturezaCarga_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NotasFiscais_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim objNatOp As New ClassNaturezaOp

    Set objNatOp = obj1

    NatOpInterna.Text = objNatOp.sCodigo

    Me.Show

    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OutrosValores_Change()
    iAlterado = REGISTRO_ALTERADO
    iOutrosAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OutrosValores_GotFocus()
    iOutrosAlterado = 0
End Sub

Private Sub OutrosValores_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OutrosValores_Validate

    If iOutrosAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(OutrosValores.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(OutrosValores.Text)
        If lErro <> SUCESSO Then gError 62886

        'Põe o valor formatado na tela
        OutrosValores.Text = Format(OutrosValores.Text, "Fixed")

    End If

    Call ValorTotal_Calcula

    Exit Sub

Erro_OutrosValores_Validate:

    Cancel = True

    Select Case gErr

        Case 62886

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Pedagio_Change()
    iAlterado = REGISTRO_ALTERADO
    iPedagioAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Pedagio_GotFocus()
    iPedagioAlterado = 0
End Sub

Private Sub Pedagio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Pedagio_Validate

    If iPedagioAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(Pedagio.ClipText)) > 0 Then
        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(Pedagio.Text)
        If lErro <> SUCESSO Then gError 62887

        'Põe o valor formatado na tela
        Pedagio.Text = Format(Pedagio.Text, "Fixed")

    End If

    Call ValorTotal_Calcula

    Exit Sub

Erro_Pedagio_Validate:

    Cancel = True


    Select Case gErr

        Case 62887

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub PedagioIncluso_Click()
    iAlterado = REGISTRO_ALTERADO
    Call ValorTotal_Calcula
End Sub

Private Sub PesoMercadoria_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PesoMercadoria_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoMercadoria_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(PesoMercadoria.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(PesoMercadoria.Text)
    If lErro <> SUCESSO Then gError 62888

    'Põe o valor formatado na tela
    PesoMercadoria.Text = Format(PesoMercadoria.Text, "Fixed")

    Exit Sub

Erro_PesoMercadoria_Validate:

    Cancel = True

    Select Case gErr

        Case 62888

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub


Private Sub SEC_Change()
    iAlterado = REGISTRO_ALTERADO
    iSECAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SEC_GotFocus()
    iSECAlterado = 0
End Sub

Private Sub SEC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SEC_Validate

    If iSECAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(SEC.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(SEC.Text)
        If lErro <> SUCESSO Then gError 62889

        'Põe o valor formatado na tela
        SEC.Text = Format(SEC.Text, "Fixed")

    End If

    Call ValorTotal_Calcula

    Exit Sub

Erro_SEC_Validate:

    Cancel = True

    Select Case gErr

        Case 62889

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Serie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Serie_Validate

    'Verifica se foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub

    'Verifica se foi selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub

    'Tenta selecionar a serie
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 62890

    'Se não está na combo
    If lErro <> SUCESSO Then

        objSerie.sSerie = Serie.Text
        'Busca a série no BD
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 62891
        If lErro <> SUCESSO Then gError 62892 'Se não encontrou

    End If

    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case gErr

        Case 62890, 62891

        Case 62892
            'Pergunta se deseja criar nova série
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
            'Se a resposta for afirmativa
            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de séries
                Call Chama_Tela("SerieNFiscal", objSerie)
            End If
            'segura o foco na série

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub SerieLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'recolhe a serie da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série na Tela
    Serie.Text = objSerie.sSerie

    Me.Show

    Exit Sub

End Sub
Public Sub NFiscalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long
Dim colSelecao As New Collection

    'Recolhe os dados da Nota Fiscal
    lErro = Move_Conhecimento_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Exit Sub

    colSelecao.Add TIPODOCINFO_CONHECIMENTOFRETE

    'Chama a Tela NFConhecimentoFreteLista
    Call Chama_Tela("NFConhecimentoFreteLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'colocado por mario em 01/10/01 pois não lia o campo de data de referencia
    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 92620
    
    If lErro <> SUCESSO Then gError 92621
   
    'Coloca na Tela a Nota Fiscal escolhida
    lErro = Traz_Conhecimento_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 62893

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 62893, 82620

        Case 92621
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Cliente.Text)) > 0 Then objCliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objNFiscal As New ClassNFiscal
Dim bCancel As Long

On Error GoTo Erro_Cliente_Validate

    'Verifica se o cliente foi alterado
    If iClienteAlterado = 0 Then Exit Sub
    'Se op cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 62894

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 62895

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
                
        If Not gbCarregandoTela Then
            'Seleciona filial na Combo Filial
            If iCodFilial = FILIAL_MATRIZ Then
                Filial.ListIndex = 0
            Else
                Call CF("Filial_Seleciona", Filial, iCodFilial)
            End If

        End If

    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        Filial.Clear
    End If

     'Limpa o GridParcelas
    ' Call Grid_Limpa(objGridParcelas)

    'Número de Parcelas

    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 62894, 62895

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 62896

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case gErr

        Case 62896

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub




Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 62897

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 62897

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 62898

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 62898

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO


End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    'Verifica se algo foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento da Filial selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 62899

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 62899

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 62950

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 62951

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 62952

        If lErro = 17660 Then gError 62953

        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        lErro = Trata_FilialCliente()
        If lErro <> SUCESSO Then gError 62954

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 62955


    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 62950, 62952

        Case 62953
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 62951
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 62954, 62955
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Carrega_PlacaUF() As Long
'Lê as Siglas dos Estados e alimenta a list da Combobox PlacaUF

Dim lErro As Long
Dim colSiglasUF As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_PlacaUF

    Set colSiglasUF = gcolUFs

    'Adiciona na Combo PlacaUF
    For iIndice = 1 To colSiglasUF.Count
        PlacaUF.AddItem colSiglasUF.Item(iIndice)
        UFRemetente.AddItem colSiglasUF.Item(iIndice)
        UFDestinatario.AddItem colSiglasUF.Item(iIndice)
    Next

    Carrega_PlacaUF = SUCESSO

    Exit Function

Erro_Carrega_PlacaUF:

    Carrega_PlacaUF = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Function

Public Sub PlacaUF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PlacaUF_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UFRemetente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UFRemetente_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(UFRemetente.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(UFRemetente)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 62956

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 62957

    Exit Sub

Erro_UFRemetente_Validate:

    Cancel = True


    Select Case gErr

        Case 62956

        Case 62957
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, UFRemetente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
Public Sub UFRemetente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UFRemetente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UFDestinatario_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UFDestinatario_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(UFDestinatario.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(UFDestinatario)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 62958

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 62959

    Exit Sub

Erro_UFDestinatario_Validate:

    Cancel = True


    Select Case gErr

        Case 62958

        Case 62959
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, UFDestinatario.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
Public Sub UFDestinatario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UFDestinatario_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 62960

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 62961

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case 62960

        Case 62961
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub


Public Sub CGCRemetente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGCRemetente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGCRemetente_Validate

    'Se CGCRemetente/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGCRemetente.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CGCRemetente.Text))

        Case STRING_CPF 'CPF

            'Critica Cpf
            lErro = Cpf_Critica(CGCRemetente.Text)
            If lErro <> SUCESSO Then gError 62962

            'Formata e coloca na Tela
            CGCRemetente.Format = "000\.000\.000-00; ; ; "
            CGCRemetente.Text = CGCRemetente.Text

        Case STRING_CGC 'CGCRemetente

            'Critica CGCRemetente
            lErro = Cgc_Critica(CGCRemetente.Text)
            If lErro <> SUCESSO Then gError 62963

            'Formata e Coloca na Tela
            CGCRemetente.Format = "00\.000\.000\/0000-00; ; ; "
            CGCRemetente.Text = CGCRemetente.Text

        Case Else

            gError 62964

    End Select

    Exit Sub

Erro_CGCRemetente_Validate:

    Cancel = True

    Select Case gErr

        Case 62962, 62963

        Case 62964
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
Public Sub CGCRemetente_GotFocus()

    Call MaskEdBox_TrataGotFocus(CGCRemetente, iAlterado)

End Sub

Public Sub CGCDestinatario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGCDestinatario_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGCDestinatario_Validate

    'Se CGCDestinatario/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGCDestinatario.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CGCDestinatario.Text))

        Case STRING_CPF 'CPF

            'Critica Cpf
            lErro = Cpf_Critica(CGCDestinatario.Text)
            If lErro <> SUCESSO Then gError 62965

            'Formata e coloca na Tela
            CGCDestinatario.Format = "000\.000\.000-00; ; ; "
            CGCDestinatario.Text = CGCDestinatario.Text

        Case STRING_CGC 'CGC

            'Critica CGCDestinatario
            lErro = Cgc_Critica(CGCDestinatario.Text)
            If lErro <> SUCESSO Then gError 62966

            'Formata e Coloca na Tela
            CGCDestinatario.Format = "00\.000\.000\/0000-00; ; ; "
            CGCDestinatario.Text = CGCDestinatario.Text

        Case Else

            gError 62967

    End Select

    Exit Sub

Erro_CGCDestinatario_Validate:

    Cancel = True


    Select Case gErr

        Case 62965, 62966

        Case 62967
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
Public Sub CGCDestinatario_GotFocus()

    Call MaskEdBox_TrataGotFocus(CGCDestinatario, iAlterado)

End Sub
Private Function Carrega_Serie() As Long
'Carrega as combos de Série e serie de NF original com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 62968

    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next

    Carrega_Serie = SUCESSO
    
    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 62968

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub BotaoLimparNF_Click()

    NFiscal.Caption = ""

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais da tela
    Set objEventoSerie = Nothing
    Set objEventoCliente = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoCondPagto = Nothing
    Set objGridParcelas = Nothing
    Set objEventoComprovServ = Nothing


'************ TRATAMENTO COMISSOES *************
    Set objGridComissoes = Nothing
    Set objTabComissoes = Nothing
'***********************************************

    Set objGrid1 = Nothing
    Set objContabil = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    'Grid ComprovServ
    Set objGridComprovServ = Nothing

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFiscal"

    'Lê os dados da Tela NFiscal
    lErro = Move_Conhecimento_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 62969

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_BUFFER_MAX_TEXTO, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_BUFFER_MAX_TEXTO, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
'    colCampoValor.Add "VolumeEspecie", objNFiscal.sVolumeEspecie, STRING_BUFFER_MAX_TEXTO, "VolumeEspecie"
'    colCampoValor.Add "VolumeMarca", objNFiscal.sVolumeMarca, STRING_BUFFER_MAX_TEXTO, "VolumeMarca"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "Observacao", objNFiscal.sObservacao, STRING_BUFFER_MAX_TEXTO, "Observacao"
    colCampoValor.Add "DataReferencia", objNFiscal.dtDataReferencia, 0, "DataReferencia"
        
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "TipoNFiscal", OP_IGUAL, TIPODOCINFO_CONHECIMENTOFRETE

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 62969

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Preenche

    Set objNFiscal.objConhecimentoFrete = New ClassConhecimentoFrete

    objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objNFiscal.lNumIntDoc <> 0 Then

        'Carrega objNFiscal com os dados passados em colCampoValor
        objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
        objNFiscal.iTipoNFiscal = colCampoValor.Item("TipoNFiscal").vValor
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.sNaturezaOp = colCampoValor.Item("NaturezaOP").vValor
        objNFiscal.lCliente = colCampoValor.Item("Cliente").vValor
        objNFiscal.iFilialCli = colCampoValor.Item("FilialCli").vValor
        objNFiscal.sSerie = colCampoValor.Item("Serie").vValor
        objNFiscal.lNumNotaFiscal = colCampoValor.Item("NumNotaFiscal").vValor
        objNFiscal.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objNFiscal.sPlaca = colCampoValor.Item("Placa").vValor
        objNFiscal.sPlacaUF = colCampoValor.Item("PlacaUF").vValor
        objNFiscal.lVolumeQuant = colCampoValor.Item("VolumeQuant").vValor
'        objNFiscal.sVolumeEspecie = colCampoValor.Item("VolumeEspecie").vValor
'        objNFiscal.sVolumeMarca = colCampoValor.Item("VolumeMarca").vValor
        objNFiscal.sVolumeNumero = colCampoValor.Item("VolumeNumero").vValor
        objNFiscal.iStatus = colCampoValor.Item("Status").vValor
        objNFiscal.sObservacao = colCampoValor.Item("Observacao").vValor
        objNFiscal.dtDataReferencia = colCampoValor.Item("DataReferencia").vValor
        
        'Coloca os dados da NFiscal na tela
        lErro = Traz_Conhecimento_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 62970

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 62970

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function Move_GridComprovServ_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados de gridComprovServ para a memória

Dim iIndice As Integer
Dim objComprovServ As ClassCompServ
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_Move_GridComprovServ_Memoria

    'Se não há comprovantes a recolher, sai da função
'    If objGridComprovServ.iLinhasExistentes = 0 Then Exit Function

    'Para cada comprovante do grid
    For iIndice = 1 To objGridComprovServ.iLinhasExistentes
                
        'Verifica se o comprovante esta relacionado com o cliente
        lErro = CF("Verifica_Cliente1", StrParaLong(GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col)), objNFiscal.lCliente)
        If lErro <> SUCESSO And lErro <> 99171 Then gError 99141
        
        'Comprovante não relacionado ao cliente
        If lErro = 99171 Then gError 99142
        
        Set objComprovServ = New ClassCompServ
        
        lErro = CF("Produto_Formata", GridComprovServ.TextMatrix(iIndice, iGrid_ServicoCon_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 99143

        'recolhe os dados do grid de comprovantes de serviços
        objComprovServ.lCodigo = StrParaLong(GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col))
        objComprovServ.dtDataEmissao = StrParaDate(GridComprovServ.TextMatrix(iIndice, iGrid_DataCon_Col))
        objComprovServ.sProduto = sProdutoFormatado
        objComprovServ.dQuantidade = StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_QuantCon_Col))
        objComprovServ.dFretePeso = StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_PrecoCon_Col))
        objComprovServ.dAdValoren = PercentParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_AdValorenCon_Col))
        objComprovServ.dValorMercadoria = StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_ValorMercadoriaCon_Col))
        objComprovServ.dValorContainer = StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_ValorContainerCon_Col))
        objComprovServ.dPedagio = StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_PedagioCon_Col))
        
        'Carrega o obj na coleção
        objNFiscal.colComprovServ.Add objComprovServ
               
    Next
    
    Move_GridComprovServ_Memoria = SUCESSO
     
    Exit Function

Erro_Move_GridComprovServ_Memoria:
    
    Move_GridComprovServ_Memoria = gErr
    
    Select Case gErr
        
        Case 99141, 99143
        
        Case 99142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLI_NAO_RELACIONADO_COMP", gErr, objNFiscal.lCliente, StrParaLong(GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col)))
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function Move_Conhecimento_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objItensNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objComprovServ As ClassCompServ
Dim iIndice As Integer

On Error GoTo Erro_Move_Conhecimento_Memoria
    
    'cyntia
    Set objNFiscal.objRastreamento = New ClassRastreamento
    
'================== DADOS IDENTIFICACAO =====================
    If Len(Trim(Cliente.ClipText)) > 0 Then
        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 62971

        'Não encontrou p Cliente --> erro
        If lErro = 12348 Then gError 62972

        objNFiscal.lCliente = objCliente.lCodigo

    End If

    'Verifica se Está Preenchido o Numero da Nota no Label
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.sSerie = Serie.Text
    objNFiscal.sNaturezaOp = NatOpInterna.Text
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.iTipoNFiscal = TIPODOCINFO_CONHECIMENTOFRETE
    objNFiscal.iTipoDocInfo = TIPODOCINFO_CONHECIMENTOFRETE
    objNFiscal.dtDataRegistro = gdtDataHoje
    objNFiscal.dtDataSaida = DATA_NULA
    objNFiscal.dtDataReferencia = StrParaDate(DataReferencia.Text)
    objNFiscal.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)

'============= DADOS COMPOSICAO FRETE ===================
    Set objNFiscal.objConhecimentoFrete = New ClassConhecimentoFrete

    With objNFiscal.objConhecimentoFrete

        .iICMSIncluso = ICMSIncluso.Value
        .dFretePeso = StrParaDbl(FretePeso.Text)
        .dFreteValor = StrParaDbl(FreteValor.Text)
        .dSEC = StrParaDbl(SEC.Text)
        .dDespacho = StrParaDbl(Despacho.Text)
        .dPedagio = StrParaDbl(Pedagio.Text)
        .dOutrosValores = StrParaDbl(OutrosValores.Text)
        .dAliquotas = PercentParaDbl(Aliquota.FormattedText)
        .dValorICMS = StrParaDbl(ValorICMS.Text)
        .dBaseCalculo = StrParaDbl(BaseCalculo.Text)
        objNFiscal.dValorTotal = StrParaDbl(ValorTotal.Caption)
        .dValorINSS = StrParaDbl(ValorINSS.Text)
        .iINSSRetido = INSSRetido.Value
        objNFiscal.dValorMercadoria = StrParaDbl(ValorMercadoria.Text)
        .iIncluiPedagio = PedagioIncluso.Value

    End With

'===================== DADOS TRANSPORTE =======================
    With objNFiscal.objConhecimentoFrete

        .sRemetente = Remetente.Text
        .sEnderecoRemetente = EnderecoRemetente.Text
        .sMunicipioRemetente = CidadeRemetente.Text
        .sUFRemetente = UFRemetente.Text
        .sCepRemetente = CEPRemetente.ClipText
        .sCGCRemetente = CGCRemetente.ClipText
        .sInscEstadualRemetente = InscEstRemetente.Text
        .sDestinatario = Destinatario.Text
        .sEnderecoDestinatario = EnderecoDestinatario.Text
        .sMunicipioDestinatario = CidadeDestinatario.Text
        .sUFDestinatario = UFDestinatario.Text
        .sCepDestinatario = CEPDestinatario.ClipText
        .sCGCDestinatario = CGCDestinatario.ClipText
        .sInscEstadualDestinatario = InscEstDestinatario.Text

    End With

'================== DADOS COMPLEMENTARES ====================

    objNFiscal.lVolumeQuant = StrParaInt(VolumeQuant.Text)
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    objNFiscal.sVolumeNumero = VolumeNumero.Text
        
    With objNFiscal.objConhecimentoFrete

        .sNaturezaCarga = NaturezaCarga.Text
        .dPesoMercadoria = StrParaDbl(PesoMercadoria.Text)
        .dValorMercadoria = StrParaDbl(ValorMercadoria.Text)
        objNFiscal.dValorMercadoria = .dValorMercadoria
        .sNotasFiscais = NotasFiscais.Text
        objNFiscal.sObservacao = Observacao.Text
        .sColeta = Coleta.Text
        .sEntrega = Entrega.Text
        .sCalculadoAte = CalculadoAte.Text
        .sMarcaVeiculo = MarcaVeiculo.Text
        .sLocalVeiculo = LocalVeiculo.Text
        objNFiscal.sPlaca = Placa.Text
        objNFiscal.sPlacaUF = PlacaUF.Text

    End With

'=========== INICIALIZANDO DADOS NAO PRESENTES NA TELA =========

    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.lNumIntDoc = 0

'    'Chama Move_GridParcelas_Memoria
'    lErro = Move_GridParcelas_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 62973

'********************* TRATAMENTO COMISSOES *************************
    'Chama Move_GridComissoes_Memoria
    lErro = objTabComissoes.Move_TabComissoes_Memoria(objNFiscal, NOTA_FISCAL)
    If lErro <> SUCESSO Then gError 42393
'********************************************************************

    'move o grid de comprovantes de serviços para a memória
    lErro = Move_GridComprovServ_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 99144
    
    For Each objComprovServ In objNFiscal.colComprovServ
    
        objProduto.sCodigo = objComprovServ.sProduto

        'Lê na Tabela de Produto, os dados do produto com o
        'código passado.
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 92974

        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 92975

        Set objItensNF = New ClassItemNF
    
        iIndice = iIndice + 1
    
        objItensNF.sProduto = objComprovServ.sProduto
        objItensNF.sUnidadeMed = objProduto.sSiglaUMVenda
        objItensNF.sUMVenda = objProduto.sSiglaUMVenda
        objItensNF.dQuantidade = objComprovServ.dQuantidade
        objItensNF.iItem = iIndice
        objItensNF.dPrecoUnitario = objComprovServ.dFretePeso
        objItensNF.dtDataEntrega = DATA_NULA
        objNFiscal.ColItensNF.Add1 objItensNF
        
    Next
    
    Move_Conhecimento_Memoria = SUCESSO

    Exit Function

Erro_Move_Conhecimento_Memoria:

    Move_Conhecimento_Memoria = gErr

    Select Case gErr

        Case 42393, 62971, 62973, 92974, 92975, 99144

        Case 62972
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 92975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Placa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
         Frame1(TabStrip1.SelectedItem.Index).Visible = True
         'Torna Frame atual visivel
         Frame1(iFrameAtual).Visible = False

        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If UCase(left(TabStrip1.SelectedItem.Caption, 6)) = UCase(left(TITULO_TAB_CONTABILIDADE, 6)) Then Call objContabil.Contabil_Carga_Modelo_Padrao
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ValorICMS_Change()

    iAlterado = REGISTRO_ALTERADO
    iICMSAlterada = REGISTRO_ALTERADO

End Sub

Private Sub ValorICMS_GotFocus()
    iICMSAlterada = 0
End Sub

Private Sub ValorICMS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Pedagio_Validate

    If iICMSAlterada = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorICMS.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorICMS.Text)
        If lErro <> SUCESSO Then gError 62974

        'Põe o valor formatado na tela
        ValorICMS.Text = Format(ValorICMS.Text, "Fixed")
    End If


    lErro = ValorTotal_Calcula(True)

    Exit Sub

Erro_Pedagio_Validate:

    Cancel = True

    Select Case gErr

        Case 62974

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ValorINSS_GotFocus()
    iValorINSSAlterado = 0
End Sub


Private Sub ValorMercadoria_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorMercadoria_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorMercadoria_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorMercadoria.Text)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(ValorMercadoria.Text)
    If lErro <> SUCESSO Then gError 62975

    'Põe o valor formatado na tela
    ValorMercadoria.Text = Format(ValorMercadoria.Text, "Fixed")

    Exit Sub

Erro_ValorMercadoria_Validate:

    Cancel = True

    Select Case gErr

        Case 62975

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub VolumeEspecie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeMarca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeQuant_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Function Traz_Conhecimento_Tela(objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim objNFiscalOriginal As New ClassNFiscal
Dim bCancel As Boolean
Dim bAlterouCT As Boolean
Dim objProduto As New ClassProduto
Dim objItemNFiscal As New ClassItemNF
Dim iIndice As Integer
Dim objTituloRec As New ClassTituloReceber

On Error GoTo Erro_Traz_Conhecimento_Tela

    'Limpa a tela
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 62976

    Set objNFiscal.objConhecimentoFrete = New ClassConhecimentoFrete
    objNFiscal.objConhecimentoFrete.lNumIntNFiscal = objNFiscal.lNumIntDoc

    lErro = ConhecimentoFrete_Le(objNFiscal.objConhecimentoFrete)
    If lErro <> SUCESSO And lErro <> 62857 Then gError 62977
    If lErro <> SUCESSO Then gError 62978

    lErro = CF("ParcelasRecNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 62979
    
     'cyntia
    'Lê a parte de Comprovante da Nota Fiscal
    lErro = CF("NFiscalComprovante_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 99094 Then gError 99128

'********** TRATAMENTO COMISSOES ***********
    'Lê as Comissões da Nota Fiscal
    lErro = CF("ComissoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 35703
'*******************************************

'============ DADOS IDENTIFICACAO =====================

    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        Status.Caption = STRING_STATUS_CANCELADO
    End If

    NatOpInterna.Text = objNFiscal.sNaturezaOp
    NFiscal.Caption = objNFiscal.lNumNotaFiscal
    Serie.Text = objNFiscal.sSerie

    'Preenche o Cliente
    Cliente.Text = objNFiscal.lCliente
    Call Cliente_Validate(bCancel)

    'Preenche a Filial do Cliente
    Call Filial_Formata(Filial, objNFiscal.iFilialCli)

    Call DateParaMasked(DataEmissao, objNFiscal.dtDataEmissao)

    '*** Traz p/ tela os dados da combo de Itens da Nota Fiscal ***


    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 84769

'    Set objItemNFiscal = objNFiscal.ColItensNF(1)

'     objItemNFiscal.iItem = 1

'     For iIndice = 0 To TipoNFiscal.ListCount - 1
'
'        If TipoNFiscal.ItemData(iIndice) = StrParaInt(objItemNFiscal.sProduto) Then
'            TipoNFiscal.Text = TipoNFiscal.List(iIndice)
'            Exit For
'        End If
'
'     Next


    '************************

'========= DADOS COMPOSICAO FRETE =============
    With objNFiscal.objConhecimentoFrete

        PedagioIncluso.Value = .iIncluiPedagio
        ICMSIncluso.Value = .iICMSIncluso
        If .dFretePeso > 0 Then FretePeso.Text = Format(.dFretePeso, "Standard")
        FreteValor.Text = Format(.dFreteValor, "Standard")
        If .dSEC > 0 Then SEC.Text = Format(.dSEC, "Standard")
        If .dDespacho > 0 Then Despacho.Text = Format(.dDespacho, "Standard")
        If .dPedagio > 0 Then Pedagio.Text = Format(.dPedagio, "Standard")
        If .dOutrosValores > 0 Then OutrosValores.Text = Format(.dOutrosValores, "Standard")
        If .dAliquotas > 0 Then Aliquota.Text = Format(.dAliquotas * 100, "Fixed")
        If .dValorICMS > 0 Then ValorICMS.Text = Format(.dValorICMS, "Standard")
        BaseCalculo.Text = Format(.dBaseCalculo, "Standard")
'Retirado por mario em 16/05/02 pois estava querendo deixar a base de calculo poder ser zerada
'        Call ValorTotal_Calcula(True)

'============= DADOS TRANSPORTE =========================

        Remetente.Text = .sRemetente
        EnderecoRemetente.Text = .sEnderecoRemetente
        CidadeRemetente.Text = .sMunicipioRemetente
        UFRemetente.Text = .sUFRemetente
        CEPRemetente.PromptInclude = False
        CEPRemetente.Text = .sCepRemetente
        CEPRemetente.PromptInclude = True
        CGCRemetente.Text = .sCGCRemetente
        Call CGCRemetente_Validate(False)
        InscEstRemetente.Text = .sInscEstadualRemetente

        Destinatario.Text = .sDestinatario
        EnderecoDestinatario.Text = .sEnderecoDestinatario
        CidadeDestinatario.Text = .sMunicipioDestinatario
        UFDestinatario.Text = .sUFDestinatario
        
        CEPDestinatario.PromptInclude = False
        CEPDestinatario.Text = .sCepDestinatario
        CEPDestinatario.PromptInclude = True

        CGCDestinatario.Text = .sCGCDestinatario
        Call CGCDestinatario_Validate(False)
        InscEstDestinatario.Text = .sInscEstadualDestinatario
    End With
'============= DADOS COMPLEMENTARES =======================

    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
'    VolumeEspecie = objNFiscal.sVolumeEspecie
'    VolumeMarca = objNFiscal.sVolumeMarca

    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a espécie dos volumes do pedido
    If objNFiscal.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objNFiscal.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    Else
        VolumeEspecie.Text = ""
    End If
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a marca dos volumes do pedido
    If objNFiscal.lVolumeMarca > 0 Then
        VolumeMarca.Text = objNFiscal.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    Else
        VolumeMarca.Text = ""
    End If
    
    VolumeNumero = objNFiscal.sVolumeNumero

    With objNFiscal.objConhecimentoFrete

        NaturezaCarga.Text = .sNaturezaCarga
        If .dPesoMercadoria > 0 Then PesoMercadoria.Text = Format(.dPesoMercadoria, "Fixed")
        If .dValorMercadoria > 0 Then ValorMercadoria.Text = Format(.dValorMercadoria, "Fixed")
        NotasFiscais.Text = .sNotasFiscais
        Observacao.Text = objNFiscal.sObservacao
        Coleta.Text = .sColeta
        Entrega.Text = .sEntrega
        CalculadoAte.Text = .sCalculadoAte
        MarcaVeiculo.Text = .sMarcaVeiculo
        Placa.Text = objNFiscal.sPlaca
        PlacaUF.Text = objNFiscal.sPlacaUF
        LocalVeiculo.Text = .sLocalVeiculo

        ValorINSS.Text = Format(.dValorINSS, "Standard")
        INSSRetido.Value = IIf(.iINSSRetido <> 0, vbChecked, vbUnchecked)

    End With

'Retirado por mario em 16/05/02 pois estava querendo deixar a base de calculo poder ser zerada
'    Call ValorTotal_Calcula
    
'    'Preenche a Condicao de Pagto
'    If objNFiscal.iCondicaoPagto > 0 Then
'        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
'        Call CondicaoPagamento_Validate(bSGECancelDummy)
'    End If

    Call DateParaMasked(DataReferencia, objNFiscal.dtDataReferencia)
    giDataReferenciaAlterada = 0
    
    'cyntia
    'Preenche o Grid com os Comprovantes da Nota Fiscal
    lErro = Preenche_GridComprovante(objNFiscal)
    If lErro <> SUCESSO Then gError 99129
    
'    Preenche o Grid de Parcelas
'    lErro = Preenche_Grid_Parcelas(objNFiscal)
'    If lErro <> SUCESSO Then gError 62980

'***************** TRATAMENTO COMISSOES ************************
    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objNFiscal)
    If lErro <> SUCESSO Then gError 39022
'****************************************************************

    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 62981

'********* por leonardo, conferir
    
    lErro = NFiscal_Le_NumIntDocCPR(objNFiscal)
    If lErro <> SUCESSO And lErro <> 84894 Then gError 84890
    
    If lErro = SUCESSO Then
    
        'carrega dados do titulo a receber associado à nota fiscal para obter a condicao de pagto
        objTituloRec.lNumIntDoc = objNFiscal.lNumIntDocCPR
        lErro = CF("TituloReceber_Le", objTituloRec, 1)
        If lErro <> SUCESSO And lErro <> 26061 Then gError 84887
        If lErro <> SUCESSO Then
        
            lErro = CF("TituloReceberBaixado_Le", objTituloRec, 1)
            If lErro <> SUCESSO And lErro <> 56570 Then gError 84888 '56572
        
        End If
    
        objNFiscal.iCondicaoPagto = objTituloRec.iCondicaoPagto
        'Preenche a Condicao de Pagto
'        If objNFiscal.iCondicaoPagto > 0 Then
'
'            CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
'            Call CondicaoPagamento_Validate(bSGECancelDummy)
'
                      
    End If
    
    ValorTotal.Caption = Format(objNFiscal.dValorTotal, "Standard")
    
    iAlterado = 0

    gbCarregandoTela = False

    Traz_Conhecimento_Tela = SUCESSO

    Exit Function

Erro_Traz_Conhecimento_Tela:

    gbCarregandoTela = False

    Traz_Conhecimento_Tela = gErr

    Select Case gErr

        Case 35703, 39022, 62976, 62977, 62979, 62980, 62981, 84887, 84888

        Case 84752, 84753, 99128, 99129
        
        Case 62978
            Call Rotina_Erro(vbOKOnly, "ERRO_CONHECIMENTOFRETE_NAO_CADASTRADO", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Rotina_Comprovantes_Tela(objComprovante As ClassCompServ) As Long
'Faz o Tratamento do Comprovante

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim dTotal As Double
Dim dTotalComp As Double
Dim dValorAdValoren As Double
Dim dValorTotAdValoren As Double
Dim dTotalPedagio As Double
Dim sProdutoEnxuto As String
Dim iLinhas As Integer
Dim iIndice2 As Integer
Dim bAchou_Comp As Boolean

On Error GoTo Erro_Rotina_Comprovantes_Tela
    
    FretePeso.Text = ""
    FreteValor.Text = ""
    Pedagio.Text = ""
    dTotal = 0
    dValorAdValoren = 0
    dValorTotAdValoren = 0
    dTotalPedagio = 0
    dTotalComp = 0
        
    For iIndice = 1 To objGridComprovServ.iLinhasExistentes
        
        objComprovante.lCodigo = StrParaLong(GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col))
        objComprovante.iFilialEmpresa = giFilialEmpresa
       
        lErro = CF("CompServGR_Le_Frete", objComprovante)
        If lErro <> SUCESSO And lErro <> 99316 Then gError 99130
        
        If lErro = 99316 Then gError 99131
                
        'Preenche os dados da mercadoria caso seja o primeiro comprovante
        If objGridComprovServ.iLinhasExistentes = 1 Then
            lErro = Preenche_Dados_Mercadoria(objComprovante)
            If lErro <> SUCESSO Then gError 99317
        End If
        
        bAchou_Comp = False
        
        For iIndice2 = 1 To objGridComprovServ.iLinhasExistentes
            If iIndice2 <> GridComprovServ.Row Then
                If objComprovante.lCodigo = GridComprovServ.TextMatrix(iIndice2, iGrid_ComprovServCon_col) Then
                    bAchou_Comp = True
                    Exit For
                End If
            End If
        Next
        
        objProduto.sCodigo = objComprovante.sProduto
            
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 99132
            
        If lErro = 28030 Then gError 99133
            
        If bAchou_Comp = False Then
            
            'coloca a mascara no produto
            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 99134
            
            'Coloca o produto já com a máscara no controle
            ServicoCon.PromptInclude = False
            ServicoCon.Text = sProdutoEnxuto
            ServicoCon.PromptInclude = True
            
            GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_ComprovServCon_col) = objComprovante.lCodigo
            GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_DataCon_Col) = Format(objComprovante.dtDataEmissao, "dd/mm/yyyy")
            GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_ServicoCon_Col) = ServicoCon.Text
            GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_DescricaoCon_Col) = objProduto.sDescricao
            GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_QuantCon_Col) = Format(objComprovante.dQuantidade, "STANDARD")
            GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_PrecoCon_Col) = Format(objComprovante.dFretePeso, "STANDARD")
            GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_PedagioCon_Col) = Format(objComprovante.dPedagio, "STANDARD")
            
            If objComprovante.dAdValoren <> 0 Then
                GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_AdValorenCon_Col) = Format(objComprovante.dAdValoren, "Percent")
                GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_ValorMercadoriaCon_Col) = Format(objComprovante.dValorMercadoria, "STANDARD")
                GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_ValorContainerCon_Col) = Format(objComprovante.dValorContainer, "STANDARD")
            End If
        End If
        
        If objComprovante.dAdValoren <> 0 Then
                
                dValorAdValoren = (objComprovante.dValorMercadoria + objComprovante.dValorContainer) * objComprovante.dAdValoren
                dValorTotAdValoren = dValorTotAdValoren + dValorAdValoren
                
        End If
        
        dTotalPedagio = dTotalPedagio + objComprovante.dPedagio
        dTotalComp = (objComprovante.dFretePeso * objComprovante.dQuantidade)
        dTotal = dTotal + dTotalComp
        
        dValorAdValoren = 0
        dTotalComp = 0
                
    Next
    
    FretePeso.Text = Format(dTotal, "STANDARD")
    FreteValor.Text = Format(dValorTotAdValoren, "STANDARD")
    Pedagio.Text = Format(dTotalPedagio, "STANDARD")
    
    dTotal = 0
    dValorTotAdValoren = 0
    dTotalPedagio = 0
    
    Call ValorTotal_Calcula
    
    Rotina_Comprovantes_Tela = SUCESSO

    Exit Function

Erro_Rotina_Comprovantes_Tela:

    Rotina_Comprovantes_Tela = gErr

    Select Case gErr
        
        Case 99130, 99132, 99317
        
        Case 99131
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objComprovante.lCodigo, objComprovante.iFilialEmpresa)
        
        Case 99133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 99134
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Function Preenche_Dados_Mercadoria(objComp As ClassCompServ) As Long

Dim objCompServItem As New ClassCompServItem
Dim lErro As Long

On Error GoTo Erro_Preenche_Dados_Mercadoria

    VolumeQuant.Text = objComp.dQuantMaterial
    VolumeEspecie.Text = objComp.sEmbalagem
    NaturezaCarga.Text = objComp.sMaterial
    ValorMercadoria.Text = objComp.dValorMercadoria
    Coleta.Text = objComp.sCidadeOrigem & "/" & objComp.sUFOrigem
    Entrega.Text = objComp.sCidadeDestino & "/" & objComp.sUFDestino
    
    lErro = CF("CompServGR_Le_CompServItem", objComp)
    If lErro <> SUCESSO Then gError 99318
    
    For Each objCompServItem In objComp.colCompServItem
        If Len(Trim(objCompServItem.sDocExtTipo)) <> 0 Then
            NotasFiscais.Text = NotasFiscais.Text & objCompServItem.sDocExtTipo & "/" & objCompServItem.sDocExtNumero & "; "
        End If
        If objCompServItem.sDocExtTipo = "CTRC" Or objCompServItem.sDocIntTipo = "CTRC" Then
            Placa.Text = objCompServItem.sPlacaCaminhao
        End If
    Next
    
    Preenche_Dados_Mercadoria = SUCESSO

    Exit Function

Erro_Preenche_Dados_Mercadoria:

    Preenche_Dados_Mercadoria = gErr

    Select Case gErr
        
        Case 99318
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFiscal Is Nothing) Then

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 62982
        If lErro <> SUCESSO Then gError 62983

        'Traz a nota para a tela
        lErro = Traz_Conhecimento_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 62984

    End If
   
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 62982, 62984

        Case 62983
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Trata_FilialCliente() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente, objCliente As New ClassCliente
Dim objTransportadora As New ClassTransportadora
Dim objNFiscal As New ClassNFiscal
Dim sCliente As String

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    sCliente = Cliente.Text
    
    'Lê a FilialCliente
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO Then gError 62985

    'Preenche endereço do destinatário e do remetente
    Call Preenche_Destinatario_Remetente(objFilialCliente)

'    lErro = Preenche_GridComprovante(objNFiscal)
'    If lErro <> SUCESSO Then gError 84761

    If ComissaoAutomatica.Value = 1 Then

        Call Grid_Limpa(objGridComissoes)

'********************** TRATAMENTO COMISSOES ********************
        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then gError 59048
'****************************************************************

    End If


    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 59048, 62985, 84761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 62986

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal1()
    If lErro <> SUCESSO Then gError 62987

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 62986, 62987

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dValorTotal As Double
Dim objNFiscal As New ClassNFiscal
Dim dFator As Double, dValorIRRF As Double
Dim vbMsgRes As VbMsgBoxResult
Dim objComprovante As New ClassCompServ

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(NatOpInterna.Text)) = 0 Then gError 62988
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 62989
    If Len(Trim(Filial.Text)) = 0 Then gError 62990
    If Len(Trim(Serie.Text)) = 0 Then gError 62991
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 62992
    If Len(Trim(FreteValor.Text)) = 0 And Len(Trim(FretePeso.Text)) = 0 Then gError 62993
'    If Len(Trim(BaseCalculo.Text)) = 0 Then gError 62994
    If Len(Trim(Remetente.Text)) = 0 Then gError 62995
    If Len(Trim(Destinatario.Text)) = 0 Then gError 62996

'    If TipoNFiscal.ListIndex = -1 Then gError 84678

    dValorTotal = StrParaDbl(ValorTotal.Caption)

    'Se o total for negativo --> Erro
    If dValorTotal < 0 Then gError 62997
        
'    If objGridComprovServ.iLinhasExistentes = 0 Then gError 99148
    
'************ TRATAMENTO COMISSOES ****************
    'Valida os dados do grid de comissões
    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError 42390
'**************************************************

     'Recolhe os dados da tela
    lErro = Move_Conhecimento_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 62998

    If Len(Trim(NFiscal.Caption)) = 0 Then
        'verifica se o cliente tem crédito.
        lErro = CF("NFiscal_Testa_Credito", objNFiscal)
        If lErro <> SUCESSO Then gError 62999
    End If

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
    If lErro <> SUCESSO Then gError 92038

    'Grava a Nota Fiscal de Saída(inclusive os dados contábeis)
    lErro = ConhecimentoFrete_Grava(objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 86000

    GL_objMDIForm.MousePointer = vbDefault

    If Len(Trim(NFiscal.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 42390, 62099, 81607, 92038, 99172

        Case 84678
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCOMPROVANTE_NAO_PREENCHIDO", gErr)

        Case 62988
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 62989
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 62990
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 62991
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 62992
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 62995
            Call Rotina_Erro(vbOKOnly, "ERRO_REMETENTE_NAO_PREENCHIDO", gErr)

        Case 62993
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORFRETE_NAO_PREENCHIDO", gErr)

        Case 62994
            Call Rotina_Erro(vbOKOnly, "ERRO_BASECALCULO_NAO_PREENCHIDA", gErr)

        Case 62996
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINATARIO_NAO_PREENCHIDO", gErr)

        Case 62997
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 62998, 62999, 86000
        
        Case 99148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_COMPROVANTE_GRID", gErr)
            
        Case 99173
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objComprovante.lCodigo, objComprovante.iFilialEmpresa)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Limpa_Tela_NFiscal

    Call Limpa_Tela(Me)

  '  Call Grid_Limpa(objGridParcelas)
'********* TRATAMENTO COMISSOES ***********
    Call Grid_Limpa(objGridComissoes)
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""
'******************************************
    Call Grid_Limpa(objGridComprovServ)

    Status.Caption = ""
    Serie.Text = ""

'    TipoNFiscal.ListIndex = -1

    NFiscal.Caption = ""
    Filial.Clear

    ValorTotal.Caption = ""
    PlacaUF.Text = ""
    UFDestinatario.Text = ""
    UFRemetente.Text = ""
    giDataReferenciaAlterada = 0
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True

    CondicaoPagamento.Text = ""
    NatOpInterna.Text = NaturezaOp_Conv34(NATUREZAOP_PADRAO_CONHECIMENTOFRETE, gdtDataAtual)

    'Incluído por Luiz Nogueira em 21/08/03
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
   'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    
    iAlterado = 0
    iClienteAlterado = 0
    
    ICMSIncluso.Value = vbChecked
    PedagioIncluso.Value = vbUnchecked

    'Fecha o Sistema de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Function

Erro_Limpa_Tela_NFiscal:

    Limpa_Tela_NFiscal = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal1() As Long
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo

Dim sNatureza As String
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal1

    sNatureza = NatOpInterna.Text

    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 86001

    NatOpInterna.Text = sNatureza

    iAlterado = 0
    
    Exit Function

Erro_Limpa_Tela_NFiscal1:

    Limpa_Tela_NFiscal1 = gErr

    Select Case gErr

        Case 86001

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 86002

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 86003

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 86002, 86003

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'Início contabilidade
Public Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Public Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

End Sub

Public Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Public Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Public Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Public Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)

End Sub

Public Sub CTBGridContabil_LeaveCell()

    Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Public Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Public Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Public Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Public Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Public Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Public Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Public Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Public Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Public Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Public Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Public Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Public Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Public Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Public Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Public Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Public Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Public Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Public Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Public Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub
Public Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

'****
Public Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

'****
Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

'****
Public Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Public Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Public Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Public Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Public Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Public Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Public Sub CTBAglutina_Click()

    Call objContabil.Contabil_Aglutina_Click

End Sub

Public Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Public Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Public Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Public Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Public Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Public Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Public Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Public Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Public Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Public Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Public Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Public Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Public Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'Traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Public Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Public Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Public Sub CTBBotaoImprimir_Click()

    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Public Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick

End Sub

Public Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Public Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click

End Sub

Public Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click

End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim sContaAux As String
Dim sContaMascarada As String
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objMnemonicoCTBValor As New ClassMnemonicoCTBValor

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        'Retorna o valor do campo Frete Peso
        Case FRETE_PESO
            'Se o campo foi preenchido
            If Len(Trim(FretePeso.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(FretePeso.Text)
            Else
                'Guarda o valor 0 na coleção
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o valor do campo Frete Valor
        Case FRETE_VALOR
            'Se o campo foi preenchido
            If Len(Trim(FreteValor.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(FreteValor.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o valor do campo Valor Total
        Case VALOR_TOTAL
            'Se o campo foi preenchido
            If Len(Trim(ValorTotal.Caption)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(ValorTotal.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o valor do campo Sec/Cat
        Case SECCAT
            'Se o campo foi preenchido
            If Len(Trim(SEC.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(SEC.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o valor do campo Despacho
        Case DESPACHO1
            'Se o campo foi preenchido
            If Len(Trim(Despacho.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(Despacho.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o valor do campo Pedágio
        Case PEDAGIO1
            'Se o campo foi preenchido
            If Len(Trim(Pedagio.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(Pedagio.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o valor do campo Outros Valores
        Case OUTROS_VALORES
            'Se o campo foi preenchido
            If Len(Trim(OutrosValores.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(OutrosValores.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o valor do campo ICMS
        Case ICMS
            'Se o campo foi preenchido
            If Len(Trim(ValorICMS.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(ValorICMS.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o valor do campo ICMS Incluso
        Case ICMS_INCLUSO
            'Guarda na coleção o valor da check box ICMSIncluso
            objMnemonicoValor.colValor.Add ICMSIncluso.Value

        'Retorna o valor do campo Base Cálculo
        Case BASE_CALCULO
            'Se o campo foi preenchido
            If Len(Trim(BaseCalculo.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(BaseCalculo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o código do Cliente selecionado
        Case CLIENTE_CODIGO
            'Se o Cliente foi selecionado
            If Len(Trim(Cliente.Text)) > 0 Then

                'Guarda no obj o parâmetro que será passado para Cliente_Le_NomeReduzido
                objCliente.sNomeReduzido = Cliente.Text

                'Lê os dados do cliente a partir do nome reduzido
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 79974

                'Guarda na coleção o código do Cliente
                objMnemonicoValor.colValor.Add objCliente.lCodigo

            'Se não selecionou o cliente
            Else
                'Adiciona o valor 0
                objMnemonicoValor.colValor.Add 0
            End If

        'Retorna o nome / razão social do Cliente selecionado
        Case CLIENTE_RAZAOSOCIAL
            'Se o Cliente foi preenchido
            If Len(Trim(Cliente.Text)) > 0 Then

                'Guarda no obj o parâmetro que será passado para Cliente_Le_NomeReduzido
                objCliente.sNomeReduzido = Cliente.Text

                'Lê os dados do cliente a partir do nome reduzido
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 79975

                'Guarda na coleção o nome /  razão social do cliente
                objMnemonicoValor.colValor.Add objCliente.sRazaoSocial

            'Se não selecionou o cliente
            Else
                'Adiciona uma string vazia
                objMnemonicoValor.colValor.Add ""
            End If

        'Tenta retornar a conta contábil de vendas do clientes, caso não encontre retorna a conta de vendas global
        Case CTA_VENDAS

                'Guarda no objCliente o parâmetro que será passado para Cliente_Le_NomeReduzido
                objCliente.sNomeReduzido = Cliente.Text

                'Lê os dados do cliente a partir do nome reduzido
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 79976

                'Guarda em objFilialCliente o parâmetro que será passado para
                objFilialCliente.lCodCliente = objCliente.lCodigo
                objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)

                'Lê os dados da Filial do cliente a partir do codigo do cliente e da filial
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO Then gError 79977

                'Se a filial do cliente possui Conta Contábil de Vendas
                If Len(objFilialCliente.sContaContabil) > 0 Then

                    'Inicializa a variável que receberá a conta
                    sContaMascarada = String(STRING_CONTA, 0)

                    'Aplica o formato à conta que foi encontrada
                    lErro = Mascara_MascararConta(objFilialCliente.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 79978

                    'Guarda a conta na coleção
                    objMnemonicoValor.colValor.Add sContaMascarada

                'Senão => procura a conta de vendas nos campos globais
                Else

                    'Guarda no obj o parâmetro que será passado para MnemonicoCTBValor_Le
                    objMnemonicoCTBValor.sMnemonico = CTA_VENDAS

                    'Lê no BD os dados do Mnemonico Global
                    lErro = CF("MnemonicoCTBValor_Le", objMnemonicoCTBValor)
                    If lErro <> SUCESSO Then gError 79979

                    'Se encontrou uma conta de global de vendas = aplica o formato e guarda a conta na coleção
                    If Len(objMnemonicoCTBValor.sValor) > 0 Then

                        'Guarda no obj a conta encontrada com seu formato
                        objMnemonicoValor.colValor.Add objMnemonicoCTBValor.sValor
                    'Se não encontrou
                    Else
                        'Guarda no obj uma string vazia
                        objMnemonicoValor.colValor.Add ""
                    End If

                End If

        'Se o mnemônico não foi tratado
        Case Else
            gError 79973

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 79974, 79975, 79976, 79977, 79978, 79979

        Case 79973
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_NF_SAIDA_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Conhecimento de Transporte"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ConhecimentoFrete"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho

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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is NatOpInterna Then
''''            Call LblNatOpInterna_Click
        ElseIf Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is ComprovServCon Then
            Call BotaoComprovante_Click
        End If

    End If

End Sub

Public Sub VolumeQuant_GotFocus()

    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)

End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    objFilial.Text = CStr(iFilial)
    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 86004

    If lErro = 17660 Then gError 86005

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 86004

        Case 86005
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

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
'***** fim do trecho a ser copiado ******


Private Function ValorTotal_Calcula(Optional bValidateICMSValor As Boolean = False) As Long
'Calcula o total de acordo com os valores informados na tela
'e a opção de ICMS Incluso ou não
Dim dFretePeso As Double
Dim dFreteValor As Double
Dim dSEC As Double
Dim dDespacho As Double
Dim dOutros As Double, dOutrosBaseIMCS As Double
Dim dValorBase As Double
Dim dAliquota As Double
Dim dValorTotal As Double
Dim dValorICMS As Double
Dim lErro As Long

Dim dValorPedagio As Double
    
On Error GoTo Erro_ValorTotal_Calcula
    
    'Recolhe os valores da tela
    dFretePeso = StrParaDbl(FretePeso.Text)
    dFreteValor = StrParaDbl(FreteValor.Text)
    dSEC = StrParaDbl(SEC.Text)
    dDespacho = StrParaDbl(Despacho.Text)
    dOutros = StrParaDbl(OutrosValores.Text)
    dAliquota = PercentParaDbl(Aliquota.FormattedText)
    dValorPedagio = StrParaDbl(Pedagio.Text)
    dValorBase = StrParaDbl(BaseCalculo.Text)
    
    dValorICMS = StrParaDbl(ValorICMS.Text)
    
    If PedagioIncluso.Value = vbUnchecked Then
        dOutrosBaseIMCS = dOutros
    Else
        dOutrosBaseIMCS = dOutros + dValorPedagio
    End If
    
    dValorTotal = dFretePeso + dFreteValor + dSEC + dDespacho + dOutros + dValorPedagio
    
    'Calcula o Subtotal sem imposto
    dValorBase = dFretePeso + dFreteValor + dSEC + dDespacho + dOutrosBaseIMCS
    
    If bValidateICMSValor Then dAliquota = 0
    
    'Se o imposto não é incluido
    If ICMSIncluso.Value = vbChecked Then
        
        'Se a aliquota estiver preenchida
        If dAliquota > 0 Or (dValorICMS > 0 And dValorBase > 0) Then
        
            If dAliquota = 0 Then
                dAliquota = dValorICMS / dValorBase
            End If
            
            'Caso tenha alguma linha do grid preenchida -> recalcula o valor do frete peso
            If objGridComprovServ.iLinhasExistentes <> 0 Then dFretePeso = Round((GridComprovServ.TextMatrix(objGridComprovServ.iLinhasExistentes, iGrid_PrecoCon_Col)) / (1 - dAliquota), 2)
            
            If objGridComprovServ.iLinhasExistentes <> 0 Then dFreteValor = Round(PercentParaDbl(GridComprovServ.TextMatrix(objGridComprovServ.iLinhasExistentes, iGrid_AdValorenCon_Col)) * ((StrParaDbl(GridComprovServ.TextMatrix(objGridComprovServ.iLinhasExistentes, iGrid_ValorContainerCon_Col)) + StrParaDbl(GridComprovServ.TextMatrix(objGridComprovServ.iLinhasExistentes, iGrid_ValorMercadoriaCon_Col))) / (1 - dAliquota)), 2)
            
            'Calcula o Subtotal sem imposto
            dValorBase = dFretePeso + dFreteValor + dSEC + dDespacho + dOutrosBaseIMCS
            dValorTotal = dFretePeso + dFreteValor + dSEC + dDespacho + dOutros + dValorPedagio
            
            'Calcula o Valor ICMS com base na aliquota informada
            dValorICMS = dValorBase * dAliquota
            
        'Se o VAlor estiver preenchido
        Else
        
            'Caso tenha alguma linha do grid preenchida -> recalcula o valor do frete peso
            If objGridComprovServ.iLinhasExistentes <> 0 Then dFretePeso = (GridComprovServ.TextMatrix(objGridComprovServ.iLinhasExistentes, iGrid_PrecoCon_Col))
            
            If objGridComprovServ.iLinhasExistentes <> 0 Then dFreteValor = PercentParaDbl(GridComprovServ.TextMatrix(objGridComprovServ.iLinhasExistentes, iGrid_AdValorenCon_Col)) * ((StrParaDbl(GridComprovServ.TextMatrix(objGridComprovServ.iLinhasExistentes, iGrid_ValorContainerCon_Col)) + StrParaDbl(GridComprovServ.TextMatrix(objGridComprovServ.iLinhasExistentes, iGrid_ValorMercadoriaCon_Col))))
            
            'Calcula o Subtotal sem imposto
            dValorBase = dFretePeso + dFreteValor + dSEC + dDespacho + dOutrosBaseIMCS
            dValorTotal = dFretePeso + dFreteValor + dSEC + dDespacho + dOutros + dValorPedagio
                    
        End If
        
    'Se o imposto é incluido
    Else
    
        'Se a aliquota estiver preenchida
        If dAliquota > 0 Then
            'Calcula o valor ICMS Com base na aliquota informada
            dValorICMS = dValorBase / ((1 / dAliquota) - 1)
            dValorICMS = Round(dValorICMS, 2)
        'Se o ValorICMS estiver preenchido
        ElseIf dValorICMS > 0 Then
            'Se o SUbtotal for positivo
            If dValorTotal > 0 Then
                'Calcula a alíquota
                dAliquota = dValorICMS / (dValorBase + dValorICMS)
            End If
        End If
        'Inclui no total o ICMS
        dValorTotal = Round(dValorTotal + dValorICMS, 2)
                
        dValorBase = dValorTotal - IIf(PedagioIncluso.Value = vbUnchecked, dValorPedagio, 0)
        
    End If
    
    If dValorICMS > dValorTotal Then gError 86006
    
    'Coloca na tela os valores calculados
    If dValorBase > 0 Then BaseCalculo.Text = Format(dValorBase, "Standard")
    If dValorICMS > 0 Then ValorICMS.Text = Format(dValorICMS, "Standard")
    If dAliquota > 0 Then Aliquota.Text = Format(dAliquota * 100, "Fixed")
    If dFretePeso > 0 Then FretePeso.Text = Format(dFretePeso, "standard")
    If dFreteValor > 0 Then FreteValor.Text = Format(dFreteValor, "standard")
    
    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    If Not gbCarregandoTela Then
        'Gera a cobranca em cima do novo valor total
       ' lErro = Cobranca_Automatica()
       'If lErro <> SUCESSO Then gError 86007
    
'****************** TRATAMENTO COMISSOES *********************
        'Faz o cálculo automático das comissões
        lErro = objTabComissoes.Comissoes_Calcula_Padrao()
        If lErro <> SUCESSO Then gError 42177
'*************************************************************
    End If
    
    ValorTotal_Calcula = SUCESSO
    
    Exit Function
    
Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr
    
    Select Case gErr
        
        Case 42177, 86007
    
        Case 86006
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORICMS_MAIOR_TOTAL", gErr, dValorICMS, dValorTotal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function
    
End Function

Function BaseCalculo_Calula_ValorTotal() As Long
'CaLcula o total de acordo com os valores informados na tela
'e a opção de ICMS Incluso ou não
Dim dFretePeso As Double
Dim dFreteValor As Double
Dim dSEC As Double
Dim dDespacho As Double
Dim dPedagio As Double
Dim dOutros As Double
Dim dValorBase As Double
Dim dAliquota As Double
Dim dValorTotal As Double
Dim dValorICMS As Double
Dim dDiferenca As Double
Dim lErro  As Long

On Error GoTo Erro_BaseCalculo_Calula_ValorTotal

    'Recolhe os valores da tela
    dFretePeso = StrParaDbl(FretePeso.Text)
    dFreteValor = StrParaDbl(FreteValor.Text)
    dSEC = StrParaDbl(SEC.Text)
    dDespacho = StrParaDbl(Despacho.Text)
    dOutros = StrParaDbl(OutrosValores.Text)
    dValorICMS = StrParaDbl(ValorICMS.Text)
    dValorBase = StrParaDbl(BaseCalculo.Text)
    dAliquota = PercentParaDbl(Aliquota.FormattedText)
    
    dValorTotal = dFretePeso + dFreteValor + dSEC + dDespacho + dOutros
    
    dPedagio = IIf(PedagioIncluso.Value, 0, StrParaDbl(Pedagio.Text))
          
    If ICMSIncluso.Value = vbUnchecked Then
        dValorTotal = dValorTotal + dValorICMS
''''        dValorICMS = dValorTotal - dValorBase - dPedagio
    Else
        dValorICMS = dAliquota * dValorBase
        dValorTotal = dValorTotal + dValorICMS
    End If
    
          
    If dValorICMS > 0 Then
        ValorICMS.Text = Format(dValorICMS, "Standard")
    Else
        ValorICMS.Text = ""
        Aliquota.Text = 0
    End If
    
    If (dValorBase - dValorTotal) > 0.0001 Then gError 86008
     
    If dValorBase <> 0 Then
        dAliquota = Round(dValorICMS, 2) / Round(dValorBase, 2)
    Else
        dAliquota = 0
    End If
    
    Aliquota.Text = Format(dAliquota * 100, "Fixed")
    ValorTotal.Caption = Format(dValorTotal, "Standard")
      
'****************** TRATAMENTO COMISSOES *********************
    'Faz o cálculo automático das comissões
    lErro = objTabComissoes.Comissoes_Calcula_Padrao()
    If lErro <> SUCESSO Then gError 42177
'*************************************************************
    
    BaseCalculo_Calula_ValorTotal = SUCESSO
    
    Exit Function
    
Erro_BaseCalculo_Calula_ValorTotal:

    BaseCalculo_Calula_ValorTotal = gErr
    
    Select Case gErr
    
        Case 42177
    
        Case 86008
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORBASE_MAIOR_SUBTOTAL", gErr, dValorBase, dValorTotal)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Function

End Function

'?????????? JA EXISTE NA TELA DE CONHECIMENTO DE FRETE SIMPLES ????????
Function ConhecimentoFrete_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'grava uma nota fiscal

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 21) As Long
Dim iIndice As Integer
Dim lErro1 As Long

On Error GoTo Erro_ConhecimentoFrete_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 62845

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 62846
    Next

    lErro = CF("NFiscal_Lock_Gravacao", alComando(), objNFiscal)
    If lErro <> SUCESSO Then gError 62847

    lErro = ConhecimentoFrete_Grava_BD(alComando, objNFiscal)
    If lErro <> SUCESSO And lErro <> 62840 Then gError 62848
    
    'Se a Nota já existe grava só a contabilidade
    If lErro = SUCESSO Then

        'verifica se o cliente possui o crédito para faturar a nota fiscal.
        'Se tiver atualiza as tabelas de cliente e estatistica de liberacao do usuario
        lErro = Processa_NFiscal_Credito(objNFiscal)
        If lErro <> SUCESSO Then gError 62849

        'Grava a Estatística do Cliente
        lErro = CF("FilialCliente_Grava_Estatistica", objNFiscal)
        If lErro <> SUCESSO Then gError 62850

        'Grava a Estatística do Produto Vendido
        lErro = CF("ProdutoVendido_Grava_Estatisticas", objNFiscal, StrParaInt(CADASTRAMENTO_DOC))
        If lErro <> SUCESSO Then gError 52965


        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then

            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError 62851

        End If
        
    Else
    
        'codigo inserido por mario em 01/10/01 para permitir alterar alguns campos de nota fiscal
        'trata da alteração dos dados da nota fiscal.
        lErro = CF("NFiscal_Alteracao", objNFiscal)
        If lErro <> SUCESSO Then gError 92618

        'codigo inserido por mario em 01/10/01 para permitir alterar alguns campos de nota fiscal
        'trata da alteração dos dados da nota fiscal.
        lErro = CF("ConhecimentoFrete_Altera", objNFiscal)
        If lErro <> SUCESSO Then gError 92630

    End If
    
    'Grava os dados contábeis (contabilidade)
    lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lCliente, objNFiscal.iFilialCli, DATA_NULA, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, NAO_AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal)
    If lErro <> SUCESSO Then gError 62852

    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 62853

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ConhecimentoFrete_Grava = SUCESSO

    Exit Function

Erro_ConhecimentoFrete_Grava:

    ConhecimentoFrete_Grava = gErr

    Select Case gErr

        Case 62845
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

       Case 84754

        Case 62846
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 52965, 62847, 62848, 62849, 62850, 62851, 62852, 92618

        Case 62853
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'?????????? JA EXISTE NA TELA DE CONHECIMENTO DE FRETE SIMPLES ????????
Function ConhecimentoFrete_Grava_BD(alComando() As Long, objNFiscal As ClassNFiscal) As Long

Dim lNumIntDoc As Long
Dim lErro As Long
Dim iClasseDocCPR As Integer
Dim lNumIntDocCPR As Long
Dim objItensNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim lNumIntItemNFiscal As Long
Dim iIndice As Integer
Dim sContaContabil As String
Dim objCompNF As New ClassCompServ
Dim lCodigo As Long
Dim lNumIntNota As Long
Dim alComando2(1 To 6) As Long
Dim colComissoesEmissao As New ColComissao


On Error GoTo Erro_ConhecimentoFrete_Grava_BD
   
    'copia alguns alComando para alComando2
    For iIndice = LBound(alComando2) To UBound(alComando2)
        alComando2(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126999
    Next
   
    If objNFiscal.lNumNotaFiscal = 0 Then

        'Gera o Número para a Nota Fiscal e atualiza a Tabela de Serie
        lErro = CF("NFiscalNumAuto", objNFiscal)
        If lErro <> SUCESSO Then gError 62835

        'verifica se a nota fiscal já está cadastrada ou se já existe uma nota com os mesmos dados em um dado periodo
        lErro = CF("NFiscal_Testa_Existencia", alComando(10), alComando(11), objNFiscal)
        If lErro <> SUCESSO And lErro <> 42417 Then gError 62836

        If lErro = 42417 Then gError 62837

    Else

        'verifica se a nota fiscal já está cadastrada ou se já existe uma nota com os mesmos dados em um dado periodo
        lErro = CF("NFiscal_Testa_Existencia", alComando(10), alComando(11), objNFiscal)
        If lErro <> SUCESSO And lErro <> 42417 Then gError 62838
        If lErro <> 42417 Then gError 62839

        'Tratar a gravação da contabilidade na rotina chamadora
        gError 62840

    End If
 
    'Busca iClasseDocCPR e lNumIntDocCPR
    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
    If lErro <> SUCESSO Then gError 62841

    objNFiscal.iClasseDocCPR = iClasseDocCPR
    objNFiscal.lNumIntDocCPR = lNumIntDocCPR

    'Obtem o Número Interno da nova Nota Fiscal
    lErro = CF("NFiscal_Automatico1", alComando(12), alComando(13), alComando(14), lNumIntDoc)
    If lErro <> SUCESSO Then gError 62842

    objNFiscal.lNumIntDoc = lNumIntDoc

    With objNFiscal

        'Insere a nova Nota Fiscal de Saida no BD
        lErro = Comando_Executar(alComando(15), "INSERT INTO NFiscal (NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, DataReferencia, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,FilialPedido, VolumeNumero, FreteRespons, Observacao, ValorMercadoria) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .dtDataReferencia, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .iFilialPedido, .sVolumeNumero, .iFreteRespons, .sObservacao, .dValorMercadoria)      'William
        If lErro <> AD_SQL_SUCESSO Then gError 62843

        objNFiscal.objConhecimentoFrete.lNumIntNFiscal = .lNumIntDoc
    End With

    With objNFiscal.objConhecimentoFrete

        lErro = Comando_Executar(alComando(16), "INSERT INTO ConhecimentoFrete (NumIntNFiscal,FretePeso,FreteValor,SEC,Despacho,Pedagio,OutrosValores,Aliquota,ValorICMS,BaseCalculo,PesoMercadoria,ValorMercadoria,NotasFiscais,Coleta,Entrega,CalculadoAte,NaturezaCarga,MarcaVeiculo,LocalVeiculo,Remetente,EnderecoRemetente,MunicipioRemetente,UFRemetente,CepRemetente,CGCRemetente,InscEstadualRemetente,Destinatario,EnderecoDestinatario,MunicipioDestinatario,UFDestinatario,CepDestinatario,CGCDestinatario,InscEstadualDestinatario,ICMSIncluso,INSSRetido,ValorINSS, IncluiPedagio) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
        .lNumIntNFiscal, .dFretePeso, .dFreteValor, .dSEC, .dDespacho, .dPedagio, .dOutrosValores, .dAliquotas, .dValorICMS, .dBaseCalculo, .dPesoMercadoria, .dValorMercadoria, .sNotasFiscais, .sColeta, .sEntrega, .sCalculadoAte, .sNaturezaCarga, .sMarcaVeiculo, .sLocalVeiculo, .sRemetente, .sEnderecoRemetente, .sMunicipioRemetente, .sUFRemetente, .sCepRemetente, .sCGCRemetente, .sInscEstadualRemetente, .sDestinatario, .sEnderecoDestinatario, .sMunicipioDestinatario, .sUFDestinatario, .sCepDestinatario, .sCGCDestinatario, .sInscEstadualDestinatario, .iICMSIncluso, .iINSSRetido, .dValorINSS, .iIncluiPedagio)
        If lErro <> AD_SQL_SUCESSO Then gError 62844

    End With

    sContaContabil = String(STRING_ITENSNF_CONTACONTABIL, 0)

    'Retira da coleção os dados referentes a Itens de NF
    For Each objItensNF In objNFiscal.ColItensNF
    
        objItensNF.lNumIntNF = objNFiscal.lNumIntDoc
        
        lErro = CF("NFiscalItem_Automatico", lNumIntItemNFiscal)
        If lErro <> SUCESSO Then gError 84773

        objItensNF.lNumIntDoc = lNumIntItemNFiscal

        With objItensNF
            lErro = Comando_Executar(alComando(19), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, Almoxarifado,Ccl, ContaContabil, PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem, ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            .lNumIntNF, .iItem, .iStatus, .sProduto, .sUMVenda, .dQuantidade, .iAlmoxarifado, .sCcl, sContaContabil, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntDocOrig)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 84770

        Exit For
    Next

    lErro = CF("ComissoesNF_Grava", alComando(17), alComando(18), objNFiscal)
    If lErro <> SUCESSO Then gError 86178
    
    'Gera as Comissões na Emissão com base nas comissões armazenadas em objNFiscal e coloca-os em colComissoesEmissao
    lErro = CF("Comissoes_Gera", alComando2(1), alComando2(2), objNFiscal, colComissoesEmissao)
    If lErro <> SUCESSO Then gError 126997

    'Grava as Comissões passadas em colComissao
    lErro = CF("Comissoes_Grava1", alComando2(3), alComando2(4), alComando2(5), alComando2(6), colComissoesEmissao)
    If lErro <> SUCESSO Then gError 126998
    
    For iIndice = LBound(alComando2) To UBound(alComando2)
        Call Comando_Fechar(alComando2(iIndice))
    Next
    
    'Para cada Comprovante da Nota Fiscal
    For Each objCompNF In objNFiscal.colComprovServ
               
        'Atualiza a tabela de comprovante com o numintnota
        lErro = Comando_ExecutarPos(alComando(20), "SELECT Codigo, NumIntNota FROM CompServGR WHERE Codigo = ?", 0, lCodigo, lNumIntNota, objCompNF.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 99149
        
        lErro = Comando_BuscarPrimeiro(alComando(20))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99150
        
        'Se não encontrou o comprovante --> erro.
        If lErro = AD_SQL_SEM_DADOS Then gError 99151
        
        'comprovante já existe em outra nota fiscal --> erro
        If lNumIntNota <> 0 Then gError 99177
        
        lErro = Comando_ExecutarPos(alComando(21), " UPDATE CompServGR SET NumIntNota = ?", alComando(20), objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 99152
        
    Next
   
    ConhecimentoFrete_Grava_BD = SUCESSO

    Exit Function

Erro_ConhecimentoFrete_Grava_BD:

    ConhecimentoFrete_Grava_BD = gErr

    Select Case gErr

        Case 44288, 62835, 62836, 62838, 62841, 62842, 86178, 84772, 84773, 126997, 126998

        Case 62837, 62839
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL_NUMAUTO", gErr)

        Case 62840

        Case 84770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSNFISCAL", gErr)

        Case 62843
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 62844
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_CONHECIMENTOFRETE", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
        
        Case 126999
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 99149, 99150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
            
        Case 99151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objCompNF.lCodigo, giFilialEmpresa)
            
        Case 99152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMPROVANTESSERVICOS", gErr)
        
        Case 99177
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_JA_FATURADO", gErr, objCompNF.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function
'?????????? JA EXISTE NA TELA DE CONHECIMENTO DE FRETE SIMPLES ????????
Function ConhecimentoFrete_Le(objConhecimentoFrete As ClassConhecimentoFrete) As Long

Dim lComando As Long
Dim lErro As Long
Dim tConhecimentoFrete As typeConhecimentoFrete

On Error GoTo Erro_ConhecimentoFrete_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 62854

    With tConhecimentoFrete
        'Inicializa as variáveis string
        .sCalculadoAte = String(STRING_CONHECIMENTOFRETE_CALCULOATE, 0)
        .sCepDestinatario = String(STRING_CEP, 0)
        .sCepRemetente = String(STRING_CEP, 0)
        .sCGCDestinatario = String(STRING_CGC, 0)
        .sCGCRemetente = String(STRING_CGC, 0)
        .sColeta = String(STRING_CONHECIMENTOFRETE_COLETA, 0)
        .sDestinatario = String(STRING_CONHECIMENTOFRETE_DESTINATARIO, 0)
        .sEnderecoDestinatario = String(STRING_ENDERECO, 0)
        .sEnderecoRemetente = String(STRING_ENDERECO, 0)
        .sEntrega = String(STRING_CONHECIMENTOFRETE_ENTREGA, 0)
        .sInscEstadualDestinatario = String(STRING_INSCR_EST, 0)
        .sInscEstadualRemetente = String(STRING_INSCR_EST, 0)
        .sLocalVeiculo = String(STRING_CIDADE, 0)
        .sMarcaVeiculo = String(STRING_CONHECIMENTOFRETE_MARCA, 0)
        .sMunicipioDestinatario = String(STRING_CIDADE, 0)
        .sMunicipioRemetente = String(STRING_CIDADE, 0)
        .sNaturezaCarga = String(STRING_CONHECIMENTOFRETE_NATUREZACARGA, 0)
        .sNotasFiscais = String(STRING_CONHECIMENTOFRETE_NOTAS, 0)
        .sRemetente = String(STRING_CONHECIMENTOFRETE_REMETENTE, 0)
        .sUFDestinatario = String(STRING_ESTADO_SIGLA, 0)
        .sUFRemetente = String(STRING_ESTADO_SIGLA, 0)

        'Busca no BD o conhecimento passado
        lErro = Comando_Executar(lComando, "SELECT FretePeso,FreteValor,SEC,Despacho,Pedagio,OutrosValores,Aliquota,ValorICMS,BaseCalculo,PesoMercadoria,ValorMercadoria,NotasFiscais,Coleta,Entrega,CalculadoAte,NaturezaCarga,MarcaVeiculo,LocalVeiculo,Remetente,EnderecoRemetente,MunicipioRemetente,UFRemetente,CepRemetente,CGCRemetente,InscEstadualRemetente,Destinatario,EnderecoDestinatario,MunicipioDestinatario,UFDestinatario,CepDestinatario,CGCDestinatario,InscEstadualDestinatario,ICMSIncluso,IncluiPedagio FROM ConhecimentoFrete WHERE NumIntNFiscal = ?", _
        .dFretePeso, .dFreteValor, .dSEC, .dDespacho, .dPedagio, .dOutrosValores, .dAliquotas, .dValorICMS, .dBaseCalculo, .dPesoMercadoria, .dValorMercadoria, .sNotasFiscais, .sColeta, .sEntrega, .sCalculadoAte, .sNaturezaCarga, .sMarcaVeiculo, .sLocalVeiculo, .sRemetente, .sEnderecoRemetente, .sMunicipioRemetente, .sUFRemetente, .sCepRemetente, .sCGCRemetente, .sInscEstadualRemetente, .sDestinatario, .sEnderecoDestinatario, .sMunicipioDestinatario, .sUFDestinatario, .sCepDestinatario, .sCGCDestinatario, .sInscEstadualDestinatario, .iICMSIncluso, .iIncluiPedagio, objConhecimentoFrete.lNumIntNFiscal)
        If lErro <> AD_SQL_SUCESSO Then gError 62855

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62856
        If lErro <> AD_SQL_SUCESSO Then gError 62857 'Não encontrou

        'Carrega o obj com os dados lidos
        objConhecimentoFrete.dAliquotas = .dAliquotas
        objConhecimentoFrete.dBaseCalculo = .dBaseCalculo
        objConhecimentoFrete.dDespacho = .dDespacho
        objConhecimentoFrete.dFretePeso = .dFretePeso
        objConhecimentoFrete.dFreteValor = .dFreteValor
        objConhecimentoFrete.dOutrosValores = .dOutrosValores
        objConhecimentoFrete.dPedagio = .dPedagio
        objConhecimentoFrete.dPesoMercadoria = .dPesoMercadoria
        objConhecimentoFrete.dSEC = .dSEC
        objConhecimentoFrete.dValorICMS = .dValorICMS
        objConhecimentoFrete.dValorMercadoria = .dValorMercadoria
        objConhecimentoFrete.iICMSIncluso = .iICMSIncluso
        objConhecimentoFrete.sCalculadoAte = .sCalculadoAte
        objConhecimentoFrete.sCepDestinatario = .sCepDestinatario
        objConhecimentoFrete.sCepRemetente = .sCepRemetente
        objConhecimentoFrete.sCGCDestinatario = .sCGCDestinatario
        objConhecimentoFrete.sCGCRemetente = .sCGCRemetente
        objConhecimentoFrete.sColeta = .sColeta
        objConhecimentoFrete.sDestinatario = .sDestinatario
        objConhecimentoFrete.sEnderecoDestinatario = .sEnderecoDestinatario
        objConhecimentoFrete.sEnderecoRemetente = .sEnderecoRemetente
        objConhecimentoFrete.sEntrega = .sEntrega
        objConhecimentoFrete.sInscEstadualDestinatario = .sInscEstadualDestinatario
        objConhecimentoFrete.sInscEstadualRemetente = .sInscEstadualRemetente
        objConhecimentoFrete.sLocalVeiculo = .sLocalVeiculo
        objConhecimentoFrete.sMarcaVeiculo = .sMarcaVeiculo
        objConhecimentoFrete.sMunicipioDestinatario = .sMunicipioDestinatario
        objConhecimentoFrete.sMunicipioRemetente = .sMunicipioRemetente
        objConhecimentoFrete.sNaturezaCarga = .sNaturezaCarga
        objConhecimentoFrete.sNotasFiscais = .sNotasFiscais
        objConhecimentoFrete.sRemetente = .sRemetente
        objConhecimentoFrete.sUFDestinatario = .sUFDestinatario
        objConhecimentoFrete.sUFRemetente = .sUFRemetente
        objConhecimentoFrete.iIncluiPedagio = .iIncluiPedagio
                       
    End With

    Call Comando_Fechar(lComando)

    ConhecimentoFrete_Le = SUCESSO

    Exit Function

Erro_ConhecimentoFrete_Le:

    ConhecimentoFrete_Le = gErr

    Select Case gErr

        Case 62854
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 62855, 62856
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONHECIMENTO_FRETE", gErr)

        Case 62857

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'????????? JÁ EXITE EM ROTINAS FAT
Friend Function Processa_NFiscal_Credito(objNFiscal As ClassNFiscal) As Long
'verifica se o cliente possui o crédito para faturar a nota fiscal.
'Se tiver atualiza as tabelas de cliente e estatistica de liberacao do usuario
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim objValorLiberadoCredito As New ClassValorLiberadoCredito
Dim lComando As Long
Dim lComando1 As Long
Dim tCliente As typeCliente
Dim sCodUsuario As String
Dim dValor As Double
Dim dtData As Date
Dim objClienteEstatistica As New ClassFilialClienteEst
Dim bNFPedido As Boolean
Dim iCreditoAprovado As Integer
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Processa_NFiscal_Credito

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44482

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 44483

    'Verifica se NFiscal é oriunda de Pedido
    If objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV Or objNFiscal.iTipoNFiscal = DOCINFO_NFISVPV Then
        bNFPedido = True
    Else
        bNFPedido = False
    End If

    'Se for testa se Pedido tem crédito aprovado
    If bNFPedido Then

        objPedidoVenda.lCodigo = objNFiscal.lNumPedidoVenda
        objPedidoVenda.iFilialEmpresa = objNFiscal.iFilialPedido

        'verifica se o pedido tem credito aprovado
        lErro = CF("BloqueiosPV_Credito_Aprovado_Testa", objPedidoVenda, iCreditoAprovado)
        If lErro <> SUCESSO Then Error 25740

    End If

    'se NF não for oriunda de PV ou se crédito não está liberado
    If (Not bNFPedido) Or iCreditoAprovado <> BLOQUEIO_CREDITO_LIBERADO Then

        'Lê os saldos e o limite de credito do Cliente
        lErro = Comando_ExecutarLockado(lComando, "SELECT LimiteCredito FROM Clientes WHERE Codigo = ?", tCliente.dLimiteCredito, objNFiscal.lCliente)
        If lErro <> AD_SQL_SUCESSO Then Error 44484

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44485

        'se não encontrou os dados do cliente
        If lErro = AD_SQL_SEM_DADOS Then Error 44486

        'loca o cliente
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 44487

        'Passa chave de objClienteEstatistica
        objClienteEstatistica.lCodCliente = objNFiscal.lCliente
        objClienteEstatistica.iFilialEmpresa = objNFiscal.iFilialEmpresa

        'Le dinamicamente o Saldo dos Titulos e dos Pedidos Liberados
        lErro = CF("Cliente_Le_Estatistica_Credito", objClienteEstatistica)
        If lErro <> SUCESSO Then Error 52955

        'Verifica se a soma dos creditos em Pedidos, Titulos e NFs ultrapassa o limite de Credito do Cliente
        If tCliente.dLimiteCredito < (objClienteEstatistica.dSaldoTitulos + objClienteEstatistica.dSaldoPedidosLiberados + objClienteEstatistica.dValorNFsNaoFaturadas + IIf(bNFPedido, 0, objNFiscal.dValorTotal)) Then

            'se um usuário não autorizou o credito ==> erro
            If Len(objNFiscal.sCodUsuario) = 0 Then Error 44488

                objLiberacaoCredito.sCodUsuario = objNFiscal.sCodUsuario

            If giTipoVersao = VERSAO_FULL Then

                'verificar se o usuário tem autorizacao para liberar o valor
                lErro = CF("LiberacaoCredito_Lock", objLiberacaoCredito)
                If lErro <> SUCESSO And lErro <> 44479 Then Error 44489

                'se não foi encontrado autorização para o usuario liberar credito
                If lErro = 44479 Then Error 44490

                'se o valor da nota ultrapassar o limite de credito que o usuario pode conceder por operacao
                If objNFiscal.dValorTotal > objLiberacaoCredito.dLimiteOperacao Then Error 44491

                objValorLiberadoCredito.sCodUsuario = objNFiscal.sCodUsuario
                objValorLiberadoCredito.iAno = Year(gdtDataAtual)

                'Lê a estatistica de liberação de credito de um usuario em um determinado ano
                lErro = CF("ValorLiberadoCredito_Lock", objValorLiberadoCredito)
                If lErro <> SUCESSO And lErro <> 44470 Then Error 44492

                'se o valor da nota ultrapassar o valor mensal que o usuario tem capacidade de liberar
                If objNFiscal.dValorTotal > objLiberacaoCredito.dLimiteMensal - objValorLiberadoCredito.adValorLiberado(Month(gdtDataAtual)) Then Error 44493

                sCodUsuario = objValorLiberadoCredito.sCodUsuario

            ElseIf giTipoVersao = VERSAO_LIGHT Then

                sCodUsuario = objNFiscal.sCodUsuario

            End If

            dValor = objNFiscal.dValorTotal
            dtData = gdtDataAtual

            'Atualiza a estatistica de liberação de credito do usuario
            lErro = CF("ValorLiberadoCredito_Grava", sCodUsuario, dValor, dtData)
            If lErro <> SUCESSO Then Error 44494

        End If

    End If

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Processa_NFiscal_Credito = SUCESSO

    Exit Function

Erro_Processa_NFiscal_Credito:

    Processa_NFiscal_Credito = Err

    Select Case Err

        Case 44482, 44483
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44484, 44485
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES1", Err, objNFiscal.lCliente)

        Case 44486
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objNFiscal.lCliente)

        Case 44487
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLIENTES", Err, objNFiscal.lCliente)

        Case 44488
           Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_CREDITO", Err, objNFiscal.lCliente)

        Case 25740, 44489, 44492, 44494, 52955

        Case 44490
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_INEXISTENTE", Err, objLiberacaoCredito.sCodUsuario)

        Case 44491
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_LIMITEOPERACAO", Err, objLiberacaoCredito.sCodUsuario)

        Case 44493
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_LIMITEMENSAL", Err, objLiberacaoCredito.sCodUsuario)

        Case 44495
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CLIENTE", Err, objNFiscal.lCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

'Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label15(Index), Source, X, Y)
'Esta fução foi comentada pois estava dando erro de compilação.
'End Sub
'
'Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
'Esta fução foi comentada pois estava dando erro de compilação.
'End Sub

Private Sub Label30_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label30(Index), Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30(Index), Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub NFiscal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscal, Source, X, Y)
End Sub

Private Sub NFiscal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscal, Button, Shift, X, Y)
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

Private Sub LblNatOpInterna_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNatOpInterna, Source, X, Y)
End Sub

Private Sub LblNatOpInterna_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNatOpInterna, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

'Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label5, Source, X, Y)
'Esta fução foi comentada pois estava dando erro de compilação.
'End Sub

'Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
'Esta fução foi comentada pois estava dando erro de compilação.
'End Sub
'
'Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label27, Source, X, Y)
'End Sub
'
'Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
'End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

'Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label9, Source, X, Y)
'Esta fução foi comentada pois estava dando erro de compilação.
'End Sub
'
'Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
'Esta fução foi comentada pois estava dando erro de compilação.
'End Sub

'Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(, Source, X, Y)
'Esta fução foi comentada pois estava dando erro de compilação.
'End Sub
'
'Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
'Esta fução foi comentada pois estava dando erro de compilação.
'End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
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

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
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

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
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

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
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

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Tratamento de saída de célula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 86009

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

'            'Se for o GridParcelas
'            Case GridParcelas.Name
'
'                lErro = Saida_Celula_GridParcelas(objGridInt)
'                If lErro <> SUCESSO Then gError 86010

'********************* TRATAMENTO COMISSOES ********************
            'Se for o GridComissoes
            Case GridComissoes.Name

                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then gError 42315
'****************************************************************
            'Se for o GridComprovServ
            Case GridComprovServ.Name

                lErro = Saida_Celula_GridComprovServ(objGridInt)
                If lErro <> SUCESSO Then gError 99153
            
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 86011

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 42315, 86009, 86010, 86011, 99153

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridComprovServ(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComprovServ

    'Verifica qual a coluna atual do Grid
    If objGridInt.objGrid.Col = iGrid_ComprovServCon_col Then
    
        lErro = Saida_Celula_ComprovServCon(objGridInt)
        If lErro <> SUCESSO Then gError 99154
        
    End If

    Saida_Celula_GridComprovServ = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComprovServ:

    Saida_Celula_GridComprovServ = gErr

    Select Case gErr

        Case 99154
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ComprovServCon(objGridComprovServ As AdmGrid) As Long
'Faz a crítica da célula do Código do Comprovante que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim objComprovante As New ClassCompServ
Dim objCliente As New ClassCliente

On Error GoTo Erro_Saida_Celula_ComprovServCon

    Set objGridComprovServ.objControle = ComprovServCon
    
    'Verifica se o comprovante foi preenchido
     If Len(Trim(ComprovServCon.ClipText)) > 0 Then
                              
        If Len(Trim(Cliente.Text)) = 0 Then gError 99162
        
        'Critica a existencia do comprovante
        objComprovante.lCodigo = StrParaLong(ComprovServCon.Text)
        objComprovante.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("CompServFrete_Le", objComprovante)
        If lErro <> SUCESSO And lErro <> 99166 Then gError 99155
        
        If lErro = 99166 Then gError 99156
        
        'comprovante já existe em outra nota fiscal --> erro
        If objComprovante.lNumIntNota <> 0 Then gError 99178
        
        objCliente.sNomeReduzido = Cliente.Text
    
        lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
        If lErro <> SUCESSO Then gError 99157
        
        If objComprovante.lCliente <> objCliente.lCodigo Then gError 99158
                              
        For iIndice = 1 To objGridComprovServ.iLinhasExistentes
            If iIndice <> GridComprovServ.Row Then
                If Trim(ComprovServCon.ClipText) = Trim(GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col)) Then gError 99159
            End If
        Next
      
        'Atualiza o total de linhas existentes no grid
        If GridComprovServ.Row - GridComprovServ.FixedRows = objGridComprovServ.iLinhasExistentes Then
            objGridComprovServ.iLinhasExistentes = objGridComprovServ.iLinhasExistentes + 1
        End If
        
        GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_ComprovServCon_col) = objComprovante.lCodigo
        
        'Faz o Tratamento do Comprovante
        lErro = Rotina_Comprovantes_Tela(objComprovante)
        If lErro <> SUCESSO Then gError 99161
      
        GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_ComprovServCon_col) = ""
      
    End If
    
    lErro = Grid_Abandona_Celula(objGridComprovServ)
    If lErro <> SUCESSO Then gError 99160
    
    Saida_Celula_ComprovServCon = SUCESSO

    Exit Function

Erro_Saida_Celula_ComprovServCon:

    Saida_Celula_ComprovServCon = gErr

    Select Case gErr

        Case 99155, 99157, 99160, 99161
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)
            
        Case 99156
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objComprovante.lCodigo, objComprovante.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)
                    
        Case 99158
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLI_NAO_RELACIONADO_COMP", gErr, objCliente.lCodigo, objComprovante.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)
            
        Case 99159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPROVANTE_JA_EXISTENTE", gErr, ComprovServCon.ClipText, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)
            
        Case 99162
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)
            
        Case 99178
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_JA_FATURADO", gErr, objComprovante.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Vencimento que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataReferencia As Date
Dim dtDataVencimento As Date
Dim sDataVencimento As String
Dim bCriouLinha As Boolean
Dim objFilialCliente As New ClassFilialCliente
Dim iContador As Integer

On Error GoTo Erro_Saida_Celula_DataVencimento

    Set objGridInt.objControle = DataVencimento

    bCriouLinha = False

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 86019

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = CDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 86020
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            bCriouLinha = True
        End If

    End If

    'If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col) Then CobrancaAutomatica.Value = vbUnchecked

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 86021

    If bCriouLinha Then
        'Coloca DescontoPadrao
'        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
'        If lErro <> SUCESSO Then gError 86022
        
        If Len(Trim(Cliente.ClipText)) > 0 Then
    
            objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
            If lErro <> SUCESSO Then gError 95212
            
            If objFilialCliente.iCodCobrador = 0 Then objFilialCliente.iCodCobrador = COBRADOR_PROPRIA_EMPRESA
        
            For iContador = 0 To Cobrador.ListCount - 1
                If Cobrador.ItemData(iContador) = objFilialCliente.iCodCobrador Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = Cobrador.List(iContador)
                    Exit For
                End If
            Next
            
'            Carrega_CarteiraCobrador (objFilialCliente.iCodCobrador)
'            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_CarteiraCobranca_Col) = CarteiraCobrador.List(0)
'
        End If
                
    End If

    Saida_Celula_DataVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr

    Select Case gErr

        Case 86019, 86021
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86022, 95212

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo de Desconto que está deixando de sser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim iTipo As Integer


On Error GoTo Erro_Saida_Celula_TipoDesconto

    If GridParcelas.Col = iGrid_Desc1Codigo_Col Then
        Set objGridInt.objControle = Desconto1Codigo
    ElseIf GridParcelas.Col = iGrid_Desc2Codigo_Col Then
        Set objGridInt.objControle = Desconto2Codigo
    ElseIf GridParcelas.Col = iGrid_Desc3Codigo_Col Then
        Set objGridInt.objControle = Desconto3Codigo
    End If

    'Verifica se o Tipo foi preenchido
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If objGridInt.objControle.Text <> objGridInt.objControle.List(objGridInt.objControle.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(objGridInt.objControle, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 86033

            'Não foi encontrado
            If lErro = 25085 Then gError 86030
            If lErro = 25086 Then gError 86031

        End If

        iTipo = Codigo_Extrai(objGridInt.objControle.Text)

        If (iTipo = VALOR_ANT_DIA) Or (iTipo = VALOR_ANT_DIA_UTIL) Or (iTipo = VALOR_FIXO) Then
            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3) = ""
        ElseIf iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual Then
            '*** Acrescentado + 1 If para contabilizar com colocação de valores de desconto
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3))) = 0 Then
                GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 2) = ""
            End If
        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    Else
        For iIndice = objGridInt.objGrid.Col To iGrid_Desc3Percentual_Col
            GridParcelas.TextMatrix(GridParcelas.Row, iIndice) = ""
        Next
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 86032

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case 86033, 86032
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoData(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Data que está deixando de sser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim dtDataVencimento As Date

On Error GoTo Erro_Saida_Celula_DescontoData

    If GridParcelas.Col = iGrid_Desc1Ate_Col Then
        Set objGridInt.objControle = Desconto1Ate
    ElseIf GridParcelas.Col = iGrid_Desc2Ate_Col Then
        Set objGridInt.objControle = Desconto2Ate
    ElseIf GridParcelas.Col = iGrid_Desc3Ate_Col Then
        Set objGridInt.objControle = Desconto3Ate
    End If

    If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then

        lErro = Data_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 86034
        'Se a data de vencimento estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) = 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If CDate(objGridInt.objControle.Text) < CDate(DataEmissao.Text) Then gError 86035
        End If

        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))) > 0 Then
            dtDataVencimento = CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))
            If CDate(objGridInt.objControle) > dtDataVencimento Then gError 86036
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 86037

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case 86034, 86037
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_INFERIOR_DATA_EMISSAO", gErr, CDate(objGridInt.objControle.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_SUPERIOR_DATA_VENCIMENTO", gErr, CDate(objGridInt.objControle.Text), dtDataVencimento)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoValor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Valor que está deixando de sser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_DescontoValor

    If GridParcelas.Col = iGrid_Desc1Valor_Col Then
        Set objGridInt.objControle = Desconto1Valor
    ElseIf GridParcelas.Col = iGrid_Desc2Valor_Col Then
        Set objGridInt.objControle = Desconto2Valor
    ElseIf GridParcelas.Col = iGrid_Desc3Valor_Col Then
        Set objGridInt.objControle = Desconto3Valor
    End If

    'Verifica se valor está preenchido
    If Len(objGridInt.objControle.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 86038

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 86039

    Saida_Celula_DescontoValor = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr

    Select Case gErr

        Case 86038, 86039
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoPerc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Percentual que está deixando de sser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String

On Error GoTo Erro_Saida_Celula_DescontoPerc

    If GridParcelas.Col = iGrid_Desc1Percentual_Col Then
        Set objGridInt.objControle = Desconto1Percentual
    ElseIf GridParcelas.Col = iGrid_Desc2Percentual_Col Then
        Set objGridInt.objControle = Desconto2Percentual
    ElseIf GridParcelas.Col = iGrid_Desc3Percentual_Col Then
        Set objGridInt.objControle = Desconto3Percentual
    End If

    'Se a Porcentagem estiver preenchida
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then
        'Critica porcentagem
        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 86040

        '***Código para colocar valores de desconto
        dPercentual = CDbl(objGridInt.objControle.Text) / 100
        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col))

        'Coloca Valor do Desconto na tela
        If dValorParcela > 0 Then
            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = sValorDesconto
        End If

    Else

        'Limpa Valor de Desconto
        GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = ""
        '***Fim Código para colocar valores de desconto

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 86041

    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case 86040, 86041
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Preenche_GridComprovante(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com as Parcelas da Nota Fiscal

Dim objComprovServ As ClassCompServ
Dim iIndice As Integer
Dim lErro As Long
Dim sProdutoEnxuto As String
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objCodigoNome As New ClassCliente
Dim objCotacaoMoeda As New ClassCotacaoMoeda
Dim dTotal As Double

On Error GoTo Erro_Preenche_GridComprovante

    'Preenche o Grid
    For Each objComprovServ In objNFiscal.colComprovServ
        
        
        iIndice = iIndice + 1
        
        objProduto.sCodigo = objComprovServ.sProduto
        
        'Usado para recolher a descrição do produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 99145
        
        If lErro = 28030 Then gError 99146
            
        'coloca a mascara no produto
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 99147
    
        'Coloca o produto já com a máscara no controle
        ServicoCon.PromptInclude = False
        ServicoCon.Text = sProdutoEnxuto
        ServicoCon.PromptInclude = True
        
        GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col) = objComprovServ.lCodigo
        GridComprovServ.TextMatrix(iIndice, iGrid_DataCon_Col) = Format(objComprovServ.dtDataEmissao, "dd/mm/yyyy")
        GridComprovServ.TextMatrix(iIndice, iGrid_ServicoCon_Col) = ServicoCon.Text
        GridComprovServ.TextMatrix(iIndice, iGrid_DescricaoCon_Col) = objProduto.sDescricao
        GridComprovServ.TextMatrix(iIndice, iGrid_QuantCon_Col) = Format(objComprovServ.dQuantidade, "STANDARD")
        GridComprovServ.TextMatrix(iIndice, iGrid_PrecoCon_Col) = Format(objComprovServ.dFretePeso, "STANDARD")
        GridComprovServ.TextMatrix(iIndice, iGrid_AdValorenCon_Col) = Format(objComprovServ.dAdValoren, "Percent")
        GridComprovServ.TextMatrix(iIndice, iGrid_ValorMercadoriaCon_Col) = Format(objComprovServ.dValorMercadoria, "STANDARD")
        GridComprovServ.TextMatrix(iIndice, iGrid_ValorContainerCon_Col) = Format(objComprovServ.dValorContainer, "STANDARD")
        GridComprovServ.TextMatrix(iIndice, iGrid_PedagioCon_Col) = Format(objComprovServ.dPedagio, "STANDARD")
        
    Next
        
    objGridComprovServ.iLinhasExistentes = iIndice

    Preenche_GridComprovante = SUCESSO

    Exit Function

Erro_Preenche_GridComprovante:

    Preenche_GridComprovante = gErr

    Select Case gErr
        
        Case 99145
        
        Case 99146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objComprovServ.sProduto)

        Case 99147
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

   End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim iTipo As Integer
Dim sCarteira As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
'        Case Desconto1Ate.Name, Desconto1Valor.Name, Desconto1Percentual.Name
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))) = 0 Then
'                objControl.Enabled = False
'            Else
'                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))
'                If objControl.Name = Desconto1Ate.Name Then
'                    objControl.Enabled = True
'                ElseIf objControl.Name = Desconto1Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
'                    Desconto1Valor.Enabled = True
'                ElseIf objControl.Name = Desconto1Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
'                    Desconto1Percentual.Enabled = True
'                Else
'                    objControl.Enabled = False
'                End If
'            End If
'
'        Case Desconto2Ate.Name, Desconto2Valor.Name, Desconto2Percentual.Name
'            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))) = 0 Then
'                objControl.Enabled = False
'            Else
'                If objControl.Name = Desconto2Ate.Name Then
'                    objControl.Enabled = True
'                ElseIf objControl.Name = Desconto2Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
'                    Desconto2Valor.Enabled = True
'                ElseIf objControl.Name = Desconto2Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
'                    Desconto2Percentual.Enabled = True
'                Else
'                    objControl.Enabled = False
'                End If
'            End If
'
'        Case Desconto3Ate.Name, Desconto3Valor.Name, Desconto3Percentual.Name
'            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))) = 0 Then
'                objControl.Enabled = False
'            Else
'                If objControl.Name = Desconto3Ate.Name Then
'                    objControl.Enabled = True
'                ElseIf objControl.Name = Desconto3Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
'                    Desconto3Valor.Enabled = True
'                ElseIf objControl.Name = Desconto3Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
'                    Desconto3Percentual.Enabled = True
'                Else
'                    objControl.Enabled = False
'                End If
'            End If
'
'        Case Desconto2Codigo.Name, Desconto3Codigo.Name
'
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, GridParcelas.Col - 4))) = 0 Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If
'
'        Case ValorParcela.Name
'            'Se o vencimento estiver preenchido, habilita o controle
'            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))) = 0 Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If
'
'
'        Case Cobrador.Name
'
'            'Verifica se a data de vencimento OU a parcela foi(-i+ram) preenchida(s)
'            If Not ((Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_col))) > 0) Or (Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))) > 0)) Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If
'
'        Case CarteiraCobrador.Name
'
'            'Guarda a Carteira que está no Grid
'            sCarteira = GridParcelas.TextMatrix(iLinha, iGrid_CarteiraCobranca_Col)
'
'            objControl.Enabled = True
'
'            CarteiraCobrador.Clear
'            'Verifica se o cobrador foi preenchido
'
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col))) > 0 Then
'                Call Carrega_CarteiraCobrador(Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col)))
'            End If
'
'            'Seleciona na Combo
'            For iIndice = 0 To CarteiraCobrador.ListCount - 1
'                If CarteiraCobrador.List(iIndice) = sCarteira Then
'                    CarteiraCobrador.ListIndex = iIndice
'                    Exit For
'                End If
'            Next
        'Cyntia
        'Código do Comprovante
        Case ComprovServCon.Name
            
            If Len(Trim(GridComprovServ.TextMatrix(GridComprovServ.Row, iGrid_ComprovServCon_col))) = 0 Then
                ComprovServCon.Enabled = True
            Else
                ComprovServCon.Enabled = False
            End If

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub BotaoConsultaTitRec_Click()
'Abre uma tela para consulta do DocCPR vinculado à nota fiscal

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoConsultaTitRec_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 86067

    'Guarda no objNFiscal os dados necessários para consultar o título
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Serie.Text
    objNFiscal.iTipoDocInfo = TIPODOCINFO_CONHECIMENTOFRETE
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dtDataEntrada = DATA_NULA

    'Guarda em objFornecedor o nome reduzido do Fornecedor
    objCliente.sNomeReduzido = Cliente.Text

    'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 86068

    'Se não encontrou o fornecedor = > erro
    If lErro = 12348 Then gError 86069

    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lCliente = objCliente.lCodigo

    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 86070

    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 86071

    Exit Sub

Erro_BotaoConsultaTitRec_Click:

    Select Case gErr

        Case 86067, 86070, 86068

        Case 86071
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA2", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)

        Case 86069
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 86072

    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 86073

    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 86074

    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Caption)) = 0 Then gError 86075

    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 86076


    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO

    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr

    Select Case gErr

        Case 86072
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 86073
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 86074
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 86075
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 86076
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function NFiscal_Le_SemNumIntDoc(objNFiscal As ClassNFiscal) As Long
'Lê os dados da nota fiscal a partir dos dados Numero, Serie, FilialEmpresa, Tipo, Fornecedor ou Cliente, FilialForn ou FilialCli, DataEmissao, DataEntrada

Dim lErro As Long
Dim lComando As Long
Dim sSelecaoSQL As String
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_SemNumIntDoc

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79711

    With tNFiscal

        'Inicializa a string que receberá a Série
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
'        .sVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
'        .sVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        .sObservacao = String(STRING_NFISCAL_OBSERVACAO, 0)
        .sCodUsuarioCancel = String(STRING_USUARIO_CODIGO, 0)
        .sMotivoCancel = String(STRING_NFISCAL_MOTIVOCANCEL, 0)

        'Define o comando SQL que será passado no select
        sSelecaoSQL = "SELECT NFiscal.NumIntDoc, NFiscal.Status, NFiscal.FilialEmpresa, NFiscal.FilialEntrega, NFiscal.DataVencimento, NFiscal.DataReferencia, NFiscal.FilialPedido, NFiscal.NumPedidoVenda, NFiscal.NumPedidoTerc, NFiscal.ClasseDocCPR, NFiscal.NumIntDocCPR, NFiscal.ValorTotal, NFiscal.ValorProdutos, NFiscal.ValorFrete, NFiscal.ValorSeguro, NFiscal.ValorOutrasDespesas, NFiscal.ValorDesconto, NFiscal.CodTransportadora, NFiscal.MensagemNota, NFiscal.TabelaPreco, NFiscal.NaturezaOp, NFiscal.PesoLiq, NFiscal.PesoBruto, NFiscal.NumIntTrib, NFiscal.Placa," & _
        "NFiscal.PlacaUF, NFiscal.VolumeQuant, NFiscal.VolumeEspecie, NFiscal.VolumeMarca, NFiscal.VolumeNumero, NFiscal.Canal, NFiscal.NumIntNotaOriginal, NFiscal.ClienteBenef, NFiscal.FilialCliBenef, NFiscal.FornecedorBenef, NFiscal.FilialFornBenef, NFiscal.FreteRespons, NFiscal.NumRecebimento, NFiscal.Observacao, NFiscal.CodUsuarioCancel, NFiscal.MotivoCancel FROM NFiscal WHERE NumNotaFiscal = ? AND FilialEmpresa = ? AND Serie = ? AND Fornecedor = ? AND Cliente = ? AND FilialForn = ? AND FilialCli = ? AND DataEmissao = ? AND DataEntrada = ? AND TipoNFiscal = ? AND Status <> ?"

        'Busca no BD os campos necessários para se definir a tela e o doc que será exibido
        lErro = Comando_Executar(lComando, sSelecaoSQL, .lNumIntDoc, .iStatus, .iFilialEmpresa, .iFilialEntrega, .dtDataVencimento, .dtDataReferencia, .iFilialPedido, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .sVolumeNumero, .iCanal, .lNumIntNotaOriginal, .lClienteBenef, .iFilialCliBenef, .lFornecedorBenef, .iFilialFornBenef, .iFreteRespons, .lNumRecebimento, .sObservacao, .sCodUsuarioCancel, .sMotivoCancel, objNFiscal.lNumNotaFiscal, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lFornecedor, objNFiscal.lCliente, objNFiscal.iFilialForn, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada, objNFiscal.iTipoDocInfo, STATUS_CANCELADO)
        If lErro <> AD_SQL_SUCESSO Then gError 79712

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79713

        'Se não encontrou a NFiscal = > erro
        If lErro = AD_SQL_SEM_DADOS Then gError 79714

    End With

    'Guarda no objNFiscal os dados obtidos no select
    Call Move_NFiscal_Obj(objNFiscal, tNFiscal)

    Call Comando_Fechar(lComando)

    NFiscal_Le_SemNumIntDoc = SUCESSO

    Exit Function

Erro_NFiscal_Le_SemNumIntDoc:

    NFiscal_Le_SemNumIntDoc = gErr

    Select Case gErr

        Case 79711
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79712, 79713
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 79714
        'Sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub Move_NFiscal_Obj(objNFiscal As ClassNFiscal, tNFiscal As typeNFiscal)

    With tNFiscal

        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.iStatus = .iStatus
        objNFiscal.iFilialEmpresa = .iFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dtDataReferencia = .dtDataReferencia
        objNFiscal.iFilialPedido = .iFilialPedido
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeQuant = .lVolumeQuant
'        objNFiscal.sVolumeEspecie = .sVolumeEspecie
'        objNFiscal.sVolumeMarca = .sVolumeMarca
        objNFiscal.sVolumeNumero = .sVolumeNumero
        objNFiscal.iCanal = .iCanal
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lClienteBenef = .lClienteBenef
        objNFiscal.iFilialCliBenef = .iFilialCliBenef
        objNFiscal.lFornecedorBenef = .lFornecedorBenef
        objNFiscal.iFilialFornBenef = .iFilialFornBenef
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.lNumRecebimento = .lNumRecebimento
        objNFiscal.sObservacao = .sObservacao
        objNFiscal.sCodUsuarioCancel = .sCodUsuarioCancel
        objNFiscal.sMotivoCancel = .sMotivoCancel
        
    End With

End Sub

Public Function NFiscal_Consulta_TituloPagar(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long

Dim objTituloPagar As New ClassTituloPagar

On Error GoTo Erro_NFiscal_Consulta_TituloPagar

    With objTituloPagar

        .lNumTitulo = objNFiscal.lNumNotaFiscal
        .lFornecedor = objNFiscal.lFornecedor
        .iFilial = objNFiscal.iFilialForn
        .dtDataEmissao = objNFiscal.dtDataEmissao
        .sSiglaDocumento = objTipoDocInfo.sTipoDocCPR

    End With

    Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)

    NFiscal_Consulta_TituloPagar = SUCESSO

    Exit Function

Erro_NFiscal_Consulta_TituloPagar:

    NFiscal_Consulta_TituloPagar = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function NFiscal_Consulta_TituloReceber(objNFiscal As ClassNFiscal) As Long

Dim objTituloReceber As New ClassTituloReceber

On Error GoTo Erro_NFiscal_Consulta_TituloReceber

    objTituloReceber.lNumIntDoc = objNFiscal.lNumIntDocCPR

    Call Chama_Tela("TituloReceber_Consulta", objTituloReceber)

    NFiscal_Consulta_TituloReceber = SUCESSO

    Exit Function

Erro_NFiscal_Consulta_TituloReceber:

    NFiscal_Consulta_TituloReceber = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function NFiscal_Consulta_NFPag(objNFiscal As ClassNFiscal) As Long

Dim objNFsPag As New ClassNFsPag

On Error GoTo Erro_NFiscal_Consulta_NFPag

    With objNFsPag

        .lNumNotaFiscal = objNFiscal.lNumNotaFiscal
        .lFornecedor = objNFiscal.lFornecedor
        .iFilial = objNFiscal.iFilialForn
        .iFilialEmpresa = objNFiscal.iFilialEmpresa
        .dtDataEmissao = objNFiscal.dtDataEmissao

    End With

    Call Chama_Tela("NFPag_Consulta", objNFsPag)

    NFiscal_Consulta_NFPag = SUCESSO

    Exit Function

Erro_NFiscal_Consulta_NFPag:

    NFiscal_Consulta_NFPag = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function NFiscal_Consulta_DebitosReceb(objNFiscal As ClassNFiscal) As Long

Dim objDebitoReceber As New ClassDebitoRecCli

On Error GoTo Erro_NFiscal_Consulta_DebitosReceb

    objDebitoReceber.lNumIntDoc = objNFiscal.lNumIntDocCPR

    Call Chama_Tela("DebitosReceb", objDebitoReceber)

    NFiscal_Consulta_DebitosReceb = SUCESSO

    Exit Function

Erro_NFiscal_Consulta_DebitosReceb:

    NFiscal_Consulta_DebitosReceb = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function NFiscal_Consulta_CreditoPagar(objNFiscal As ClassNFiscal) As Long

Dim objCreditoPagar As New ClassCreditoPagar

On Error GoTo Erro_NFiscal_Consulta_CreditoPagar

    objCreditoPagar.lNumIntDoc = objNFiscal.lNumIntDocCPR

    Call Chama_Tela("CreditoPagar", objCreditoPagar)

    NFiscal_Consulta_CreditoPagar = SUCESSO

    Exit Function

Erro_NFiscal_Consulta_CreditoPagar:

    NFiscal_Consulta_CreditoPagar = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

'*** Fim do trecho que deverá ser apagado ***
Private Sub Limpa_Remetente()

    Remetente.Text = ""
    EnderecoRemetente.Text = ""
    CidadeRemetente.Text = ""
    UFRemetente.Text = ""
    CEPRemetente.PromptInclude = False
    CEPRemetente.Text = ""
    CEPRemetente.PromptInclude = True
    CGCRemetente.PromptInclude = False
    CGCRemetente.Text = ""
    CGCRemetente.PromptInclude = True
    InscEstRemetente.PromptInclude = False
    InscEstRemetente.Text = ""
    InscEstRemetente.PromptInclude = True

End Sub

Private Sub Limpa_Destinatario()

    Destinatario.Text = ""
    EnderecoDestinatario.Text = ""
    CidadeDestinatario.Text = ""
    UFDestinatario.Text = ""
    CEPDestinatario.PromptInclude = False
    CEPDestinatario.Text = ""
    CEPDestinatario.PromptInclude = True
    CGCDestinatario.PromptInclude = False
    CGCDestinatario.Text = ""
    CGCDestinatario.PromptInclude = True
    InscEstDestinatario.PromptInclude = False
    InscEstDestinatario.Text = ""
    InscEstDestinatario.PromptInclude = True

End Sub

Public Function Preenche_Destinatario_Remetente(objFilialCliente As ClassFilialCliente) As Long
'Preenche o endereço do remetente e destinatario

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim objCliente As New ClassCliente

On Error GoTo Erro_Preenche_Destinatario_Remetente

    'Atribui o código do cliente ao objCliente
    objCliente.lCodigo = objFilialCliente.lCodCliente

    'Faz leitura na tabela Cliente afim de extrair o nome
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 87330

    'Se não achou cliente então fornece erro
    If lErro = 12293 Then gError 87331

    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO Then gError 87332

    'Verifica se possui endereço
    If objFilialCliente.lEnderecoEntrega > 0 Then

        'Atribui código do endereco de entrega ao objEndereço
        objEndereco.lCodigo = objFilialCliente.lEnderecoEntrega

        'Faz leitura na tabela de endereços
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 87334

        'Se não achou endereço - erro
        If lErro = 12309 Then gError 87335

        If Len(Trim(objEndereco.sEndereco)) > 0 Then

            'Traz os dados do destinatário para tela
            Call Traz_Destinatario_Tela(objEndereco, objCliente)

            'Traz o dados do remetente para a tela
            Call Traz_Remetente_Tela(objEndereco, objCliente)

        Else

            If objFilialCliente.lEndereco > 0 Then

                'Atribui código do endereco Principal ao objEndereço
                objEndereco.lCodigo = objFilialCliente.lEndereco

                'Le o endereço com o código passado
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO And lErro <> 12309 Then gError 87336

                'Se não achou endereço - erro
                If lErro = 12309 Then gError 87337

                'Traz os dados do destinatário para tela
                Call Traz_Destinatario_Tela(objEndereco, objCliente)

                'Traz o dados do remetente para a tela
                Call Traz_Remetente_Tela(objEndereco, objCliente)

            End If
        End If
    End If



Preenche_Destinatario_Remetente = SUCESSO

    Exit Function

Erro_Preenche_Destinatario_Remetente:

    Preenche_Destinatario_Remetente = gErr

    Select Case gErr

        Case 87330, 87332, 87334, 87336
        'Erros tratado na rotina

        Case 87331, 87335, 87337
        'não fornecer mensagem ao usuário, não é obrigatório
        'a existencia de um endereço cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

        End Select

    Exit Function

End Function


Public Function Traz_Destinatario_Tela(objEndereco As ClassEndereco, objCliente As ClassCliente) As Long

    Call Limpa_Destinatario

    With objEndereco

        If Len(Trim(.sEndereco)) > 0 Then
            EnderecoDestinatario.Text = .sEndereco
        End If

        If Len(Trim(.sCidade)) > 0 Then
            CidadeDestinatario.Text = .sCidade
        End If

        If Len(Trim(.sSiglaEstado)) > 0 Then
            UFDestinatario.Text = .sSiglaEstado
        End If

        If Len(Trim(.sCEP)) > 0 Then
            CEPDestinatario.PromptInclude = False
            CEPDestinatario.Text = .sCEP
            CEPDestinatario.PromptInclude = True
        End If

    End With

    With objCliente

        Destinatario.Text = .sRazaoSocial

        If Len(Trim(.sCGC)) > 0 Then
            CGCDestinatario.Text = .sCGC
            Call CGCDestinatario_Validate(bSGECancelDummy)
        End If

        If Len(Trim(.sInscricaoEstadual)) > 0 Then
            InscEstDestinatario.PromptInclude = False
            InscEstDestinatario.Text = .sInscricaoEstadual
            InscEstDestinatario.PromptInclude = True
        End If

    End With

End Function

Public Function Traz_Remetente_Tela(objEndereco As ClassEndereco, objCliente As ClassCliente) As Long

    Call Limpa_Remetente

    With objEndereco

        If Len(Trim(.sEndereco)) > 0 Then
            EnderecoRemetente.Text = .sEndereco
        End If

        If Len(Trim(.sCidade)) > 0 Then
            CidadeRemetente.Text = .sCidade
        End If

        If Len(Trim(.sSiglaEstado)) > 0 Then
            UFRemetente.Text = .sSiglaEstado
        End If

        If Len(Trim(.sCEP)) > 0 Then
            CEPRemetente.PromptInclude = False
            CEPRemetente.Text = .sCEP
            CEPRemetente.PromptInclude = True
        End If

    End With

    With objCliente

        Remetente.Text = .sRazaoSocial

        If Len(Trim(.sCGC)) > 0 Then
            CGCRemetente.Text = .sCGC
            Call CGCRemetente_Validate(bSGECancelDummy)

        End If

        If Len(Trim(.sInscricaoEstadual)) > 0 Then
            InscEstRemetente.PromptInclude = False
            InscEstRemetente.Text = .sInscricaoEstadual
            InscEstRemetente.PromptInclude = True
        End If

    End With

End Function

Private Sub ValorINSS_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorINSSAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorINSS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_ValorINSS_Validate

    If iValorINSSAlterado <> REGISTRO_ALTERADO Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorINSS.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorINSS.Text)
        If lErro <> SUCESSO Then Error 26144

        'Põe o valor formatado na tela
        ValorINSS.Text = Format(ValorINSS.Text, "Standard")

    End If

    'If INSSRetido.Value = vbChecked Then Call Cobranca_Automatica

    Exit Sub

Erro_ValorINSS_Validate:

    Cancel = True

    Select Case Err

        Case 26144, 49752

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub

End Sub

Private Sub INSSRetido_Click()

On Error GoTo Erro_INSSRetido_Click

    If gbCarregandoTela Then Exit Sub

    iAlterado = REGISTRO_ALTERADO

    'If Len(Trim(ValorINSS)) <> 0 Then Call Cobranca_Automatica

    Exit Sub

Erro_INSSRetido_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub


'***************** TRATAMENTO COMISSOES ********************
Public Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridComissoes)
    'Se alguma comissao foi excluída
    If KeyCode = vbKeyDelete Then
        'atualiza os painéis totalizadores
        Call objTabComissoes.Soma_Percentual
        Call objTabComissoes.Soma_Valor
    End If

End Sub
Public Sub ComissaoAutomatica_Click()

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Click

    iAlterado = REGISTRO_ALTERADO

    'Se a comissão automática estiver selecionada
    If ComissaoAutomatica.Value = vbChecked Then
        'Recalcula as comissoes
        lErro = objTabComissoes.Comissoes_Calcula_Padrao()
        If lErro <> SUCESSO Then gError 51616

    End If

    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr

        Case 51616

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
End Sub
Public Sub PercentualComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub PercentualComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub PercentualComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub PercentualComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentualEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub PercentualEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub PercentualEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub PercentualEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Public Sub ValorBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub ValorBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub ValorComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub ValorComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub ValorEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub ValorEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Vendedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub Vendedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Public Sub BotaoVendedores_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVendedores_Click

    'Chama a tela de browse de Vendedores
    lErro = objTabComissoes.BotaoVendedores_Click()
    If lErro <> SUCESSO Then gError 43696

    Exit Sub

Erro_BotaoVendedores_Click:

    Select Case gErr

        Case 43696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim lErro As Long

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1

    If GridComissoes.Row > 0 Then

        'Preenche a linha de Vendedor com dados default
        lErro = objTabComissoes.VendedorLinha_Preenche(objVendedor)
        If lErro <> SUCESSO Then gError 51617

    End If


    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 51617  'tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub
Public Sub GridComissoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Public Sub GridComissoes_EnterCell()

    Call Grid_Entrada_Celula(objGridComissoes, iAlterado)

End Sub

Public Sub GridComissoes_GotFocus()

    Call Grid_Recebe_Foco(objGridComissoes)

End Sub

Public Sub GridComissoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComissoes, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Public Sub GridComissoes_LeaveCell()

    Call Saida_Celula(objGridComissoes)

End Sub

Public Sub GridComissoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComissoes)

End Sub

Public Sub GridComissoes_RowColChange()

    Call Grid_RowColChange(objGridComissoes)

End Sub

Public Sub GridComissoes_Scroll()

    Call Grid_Scroll(objGridComissoes)

End Sub
'*************************************************************

 'Cyntia
Function Inicializa_Grid_ComprovServ(objGridInt As AdmGrid) As Long
'Inicializa o Grid de ComprovServ

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Comprovante")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Descrição do Serviço")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("Ad Valoren")
    objGridInt.colColuna.Add ("Pedágio")
    objGridInt.colColuna.Add ("Valor Mercadoria")
    objGridInt.colColuna.Add ("Valor Container")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (ComprovServCon.Name)
    objGridInt.colCampo.Add (DataCon.Name)
    objGridInt.colCampo.Add (ServicoCon.Name)
    objGridInt.colCampo.Add (DescricaoCon.Name)
    objGridInt.colCampo.Add (QuantCon.Name)
    objGridInt.colCampo.Add (PrecoCon.Name)
    objGridInt.colCampo.Add (AdValorenCon.Name)
    objGridInt.colCampo.Add (PedagioCon.Name)
    objGridInt.colCampo.Add (ValorMercadoriaCon.Name)
    objGridInt.colCampo.Add (ValorContainerCon.Name)
    
    'Colunas do Grid
    iGrid_ComprovServCon_col = 1
    iGrid_DataCon_Col = 2
    iGrid_ServicoCon_Col = 3
    iGrid_DescricaoCon_Col = 4
    iGrid_QuantCon_Col = 5
    iGrid_PrecoCon_Col = 6
    iGrid_AdValorenCon_Col = 7
    iGrid_PedagioCon_Col = 8
    iGrid_ValorMercadoriaCon_Col = 9
    iGrid_ValorContainerCon_Col = 10
   
    'Grid do GridInterno
    objGridInt.objGrid = GridComprovServ

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_CONHECFRETE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridComprovServ.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Usado para que se possa utilizar a Rotina_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ComprovServ = SUCESSO

    Exit Function

End Function

Private Sub Soma_Coluna()

Dim iIndice As Integer
Dim dValorMerc As Double
Dim dValorContainer As Double
Dim bValidateICMS As Boolean
Dim dPesoFrete As Double
Dim dAdValoren As Double
Dim dOutrosValores As Double
Dim dValorISS As Double
Dim dPedagio As Double
Dim dDespacho As Double
Dim dManuseio As Double
Dim lErro As Long

On Error GoTo Erro_Soma_Coluna

    iIndice = 0

    'Para o nº de colunas existentes do grid
    For iIndice = 1 To objGridComprovServ.iLinhasExistentes
        
        'Acumula nas variáveis, a soma dos valores selecionados
        'em cada coluna
        dValorContainer = dValorContainer + StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_ValorContainerCon_Col))
        dPesoFrete = dPesoFrete + StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_PrecoCon_Col))
        dAdValoren = dAdValoren + StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_AdValorenCon_Col))
        dValorMerc = dValorMerc + StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_ValorMercadoriaCon_Col))
        dPedagio = dPedagio + StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_PedagioCon_Col))
        
        
    Next

    'Joga nos controles da tela os valores da soma de cada coluna
    ValorMercadoria.Text = Format(dValorMerc, "STANDARD")
    FretePeso.Text = Format(dPesoFrete + dValorISS, "STANDARD")
    FreteValor.Text = Format(dAdValoren, "STANDARD")
    Pedagio.Text = Format(dPedagio, "STANDARD")
    
    bValidateICMS = True

    lErro = ValorTotal_Calcula(bValidateICMS)
    If lErro <> SUCESSO Then gError 84774

    Exit Sub

Erro_Soma_Coluna:

    Select Case gErr

        Case 84774

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
'*************************************************************

'Private Function Carrega_TipoNFiscal() As Long
''Carrega as combos de Série e serie de NF original com as séries lidas do BD
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim objNFiscal As New ClassNFiscal
'On Error GoTo Erro_Carrega_TipoNFiscal
'
'    'Carrega na Combo o 1º elemento
'    objProduto.sCodigo = "1001"
'
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 84763
'
'    TipoNFiscal.AddItem objProduto.sNomeReduzido
'    TipoNFiscal.ItemData(0) = 1001
'
'    'Carrega na Combo o 2º elemento
'    objProduto.sCodigo = "1002"
'
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 84764
'
'    TipoNFiscal.AddItem objProduto.sNomeReduzido
'    TipoNFiscal.ItemData(1) = 1002
'
'    'Carrega na Combo o 3º elemento
'    objProduto.sCodigo = "1008"
'
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 84765
'
'    TipoNFiscal.AddItem objProduto.sNomeReduzido
'    TipoNFiscal.ItemData(2) = 1008
'
'
'    Carrega_TipoNFiscal = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_TipoNFiscal:
'
'    Carrega_TipoNFiscal = gErr
'
'    Select Case gErr
'
'        Case 84763, 84764, 84765
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function


Public Function NFiscal_Le_NumIntDocCPR(objNFiscal As ClassNFiscal) As Long
'Lê os dados da nota fiscal a partir dos dados Numero, Serie, FilialEmpresa, Tipo, Fornecedor ou Cliente, FilialForn ou FilialCli, DataEmissao, DataEntrada

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDocCPR As Integer
On Error GoTo Erro_NFiscal_Le_NumIntDocCPR

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 84891

    lErro = Comando_Executar(lComando, "SELECT NumIntDocCPR FROM NFiscal WHERE NumIntDoc = ?", lNumIntDocCPR, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 84892
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84893
    
    If lErro = AD_SQL_SEM_DADOS Then gError 84894
    
    objNFiscal.lNumIntDocCPR = lNumIntDocCPR
    
    Call Comando_Fechar(lComando)

    NFiscal_Le_NumIntDocCPR = SUCESSO

    Exit Function

Erro_NFiscal_Le_NumIntDocCPR:

    NFiscal_Le_NumIntDocCPR = gErr

    Select Case gErr

        Case 84891
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 84894
        'Sem dados

        Case 84892, 84893
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr, Error$)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function ConhecFreteFat_Atualiza_Frete() As Long

'???????????? Remover
'Função feita p/ GSilva, serve p/ criar Itens de Notas Fiscais p/ os
'conhecimentos de frete.

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim sVolumeEspecie As String
Dim lNumIntDoc As Long
Dim objItensNFiscal As New ClassItemNF
Dim lNumIntItemNFiscal As Long
Dim dValorTotal As Double
Dim objProduto As New ClassProduto
Dim lErro2 As Long
Dim objNFiscal As New ClassNFiscal
Dim sContaContabil As String
Dim lTransacao As Long
Dim iTipoNFiscal As Integer
Dim iStatus As Integer

On Error GoTo Erro_ConhecFreteFat_Atualiza_Frete

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 10 '???
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 1 '???
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 2 '???

    sVolumeEspecie = String(20, 0)
    
    lErro = Comando_Executar(lComando, "SELECT VolumeEspecie, NumIntDoc, ValorTotal, TipoNFiscal, Status FROM NFiscal WHERE TipoNFiscal = 116 AND NOT EXISTS (SELECT * FROM ItensNFiscal WHERE ItensNFiscal.NumIntNF = NFiscal.NumIntDoc)", sVolumeEspecie, lNumIntDoc, dValorTotal, iTipoNFiscal, iStatus)
    If lErro <> AD_SQL_SUCESSO Then gError 3 '???
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 4 '???

    Do While lErro = AD_SQL_SUCESSO
        
        objNFiscal.iStatus = iStatus
        objNFiscal.iTipoNFiscal = iTipoNFiscal
        objNFiscal.iTipoNFiscal = 116
'        objNFiscal.sVolumeEspecie = sVolumeEspecie
        objNFiscal.lNumIntDoc = lNumIntDoc
        objNFiscal.dValorTotal = dValorTotal
        
        objItensNFiscal.lNumIntNF = objNFiscal.lNumIntDoc
        objItensNFiscal.iItem = 1
        
        objItensNFiscal.dQuantidade = 1

        'Se as primeiras 4 letras de VolumeEspecie for CNTR,
        'o código do produto será 1002, se não 1001.
'        If UCase(Left(objNFiscal.sVolumeEspecie, 4)) = "CNTR" Or UCase(Left(objNFiscal.sVolumeEspecie, 9)) = "CONTAINER" Then
'            objProduto.sCodigo = "1002"
'        Else
'            objProduto.sCodigo = "1001"
'        End If
        
        'Lê na Tabela de Produto, os dados do produto com o
        'código passado.
        lErro2 = CF("Produto_Le", objProduto)
        If lErro2 <> SUCESSO And lErro <> 28030 Then gError 6 '???
        
        objItensNFiscal.sUnidadeMed = objProduto.sSiglaUMVenda
        objItensNFiscal.sUMVenda = objProduto.sSiglaUMVenda
        objItensNFiscal.dQuantidade = 1
        objItensNFiscal.iItem = 1
        objItensNFiscal.dPrecoUnitario = objNFiscal.dValorTotal
        objItensNFiscal.sProduto = objProduto.sCodigo
        objItensNFiscal.sDescricaoItem = objProduto.sDescricao
        objItensNFiscal.dtDataEntrega = DATA_NULA
        objItensNFiscal.iClasseUM = objProduto.iClasseUM
        
        objItensNFiscal.lNumIntNF = objNFiscal.lNumIntDoc
        
        'Gera um lNumIntDoc para o ItemNFiscal
        lErro2 = CF("NFiscalItem_Automatico", lNumIntItemNFiscal)
        If lErro2 <> SUCESSO Then gError 7 '???

        objItensNFiscal.lNumIntDoc = lNumIntItemNFiscal
        
        'Insere itens em ItemNFiscal
        With objItensNFiscal
            lErro2 = Comando_Executar(lComando2, "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, Almoxarifado,Ccl, ContaContabil, PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem, ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            .lNumIntNF, .iItem, .iStatus, .sProduto, .sUMVenda, .dQuantidade, .iAlmoxarifado, .sCcl, sContaContabil, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntDocOrig)
        End With
        If lErro2 <> AD_SQL_SUCESSO Then gError 8 '???

        'Verifica se não é uma nota cancelada
        If objNFiscal.iStatus <> STATUS_CANCELADO Then
        
            'Grava a Estatística do Produto Vendido
            lErro2 = CF("ProdutoVendido_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_DOC)
            If lErro2 <> SUCESSO Then gError 9 '???
    
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9 '???
    
    Loop
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 11 '???
  
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    lErro = Rotina_Aviso(vbOKOnly, "A criação dos Itens das Notas Fiscais foi realizada com Sucesso!!!")
      
    ConhecFreteFat_Atualiza_Frete = SUCESSO

    Exit Function

Erro_ConhecFreteFat_Atualiza_Frete:
    
    ConhecFreteFat_Atualiza_Frete = gErr
    
    Select Case gErr
    
        Case 1, 2
            lErro = Rotina_Erro(vbOKOnly, "Erro na Abertura de Comando.", gErr)
            
        Case 3, 4, 9
            lErro = Rotina_Erro(vbOKOnly, "Erro na leitura da tabela de Notas Fiscais", gErr)
        
        Case 5, 6, 7
        
        Case 8
            lErro = Rotina_Erro(vbOKOnly, "Erro na Inserção da tabela de Itens de Notas Fiscais", gErr)

        Case 10
            lErro = Rotina_Erro(vbOKOnly, "Erro na abertura de Transacao.", gErr)
        
       Case 11
            lErro = Rotina_Erro(vbOKOnly, "Erro no Commit da Transacao.", gErr)
        
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select

    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

'????? FERNANDO
'******************JÁ EXISTE NA TELA DE COTACAOMOEDA *****************
Public Function CotacaoMoeda_Le(objCotacaoMoeda As ClassCotacaoMoeda) As Long

Dim lComando As Long
Dim lErro As Long
Dim tCotacaoMoeda As typeCotacaoMoeda

On Error GoTo Erro_CotacaoMoeda_Le

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80264

    'Verifica se existe cotação para a data informada.
    lErro = Comando_Executar(lComando, "SELECT Valor FROM CotacoesMoeda WHERE Data = ? AND Moeda = ?", tCotacaoMoeda.dValor, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
    If lErro <> AD_SQL_SUCESSO Then gError 80265

    '===> Caso exista seleciona
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80266

    '===> Se não exite Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 80267

    'Carrega objCotacaoMoeda
    objCotacaoMoeda.dValor = tCotacaoMoeda.dValor
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    CotacaoMoeda_Le = SUCESSO

Exit Function

Erro_CotacaoMoeda_Le:

    CotacaoMoeda_Le = gErr

    Select Case gErr

        Case 80264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80265, 80266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA", gErr, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)

        Case 80267 'CotacaoMoeda não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ConhecFreteFat_Atualiza_Frete2() As Long
'???????????? Remover
'Função feita p/ GSilva, serve p/ atualizar as estatísticas p/ os
'conhecimentos de frete.

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim sVolumeEspecie As String
Dim lNumIntDoc As Long
Dim objItensNFiscal As New ClassItemNF
Dim lNumIntItemNFiscal As Long
Dim dValorTotal As Double
Dim objProduto As New ClassProduto
Dim lErro2 As Long
Dim objNFiscal As New ClassNFiscal
Dim sContaContabil As String
Dim lTransacao As Long
Dim iTipoNFiscal As Integer
Dim iStatus As Integer
Dim dtDataRef As Date

On Error GoTo Erro_ConhecFreteFat_Atualiza_Frete

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 10 '???
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 1 '???
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 2 '???

    sVolumeEspecie = String(20, 0)
    
    dtDataRef = CDate("06/07/2001")
    
    lErro = Comando_Executar(lComando, "SELECT VolumeEspecie, NumIntDoc, ValorTotal, TipoNFiscal, Status FROM NFiscalConhecimentoFrete WHERE DataReferencia <= ? AND Status <> ?", sVolumeEspecie, lNumIntDoc, dValorTotal, iTipoNFiscal, iStatus, dtDataRef, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 3
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 4 '???

    Do While lErro = AD_SQL_SUCESSO
        
        objNFiscal.iStatus = iStatus
        objNFiscal.iTipoNFiscal = iTipoNFiscal
        objNFiscal.iTipoNFiscal = 116
'        objNFiscal.sVolumeEspecie = sVolumeEspecie
        objNFiscal.lNumIntDoc = lNumIntDoc
        objNFiscal.dValorTotal = dValorTotal
        
        objItensNFiscal.lNumIntNF = objNFiscal.lNumIntDoc
        objItensNFiscal.iItem = 1
        
        objItensNFiscal.dQuantidade = 1

'        'Se as primeiras 4 letras de VolumeEspecie for CNTR,
'        'o código do produto será 1002, se não 1001.
'        If UCase(Left(objNFiscal.sVolumeEspecie, 4)) = "CNTR" Or UCase(Left(objNFiscal.sVolumeEspecie, 9)) = "CONTAINER" Then
'            objProduto.sCodigo = "1002"
'        Else
'            objProduto.sCodigo = "1001"
'        End If
        
        'Lê na Tabela de Produto, os dados do produto com o
        'código passado.
        lErro2 = CF("Produto_Le", objProduto)
        If lErro2 <> SUCESSO And lErro <> 28030 Then gError 6 '???
        
        objItensNFiscal.sUnidadeMed = objProduto.sSiglaUMVenda
        objItensNFiscal.sUMVenda = objProduto.sSiglaUMVenda
        objItensNFiscal.dQuantidade = 1
        objItensNFiscal.iItem = 1
        objItensNFiscal.dPrecoUnitario = objNFiscal.dValorTotal
        objItensNFiscal.sProduto = objProduto.sCodigo
        objItensNFiscal.sDescricaoItem = objProduto.sDescricao
        objItensNFiscal.dtDataEntrega = DATA_NULA
        objItensNFiscal.iClasseUM = objProduto.iClasseUM
        
        objItensNFiscal.lNumIntNF = objNFiscal.lNumIntDoc
        
'        'Gera um lNumIntDoc para o ItemNFiscal
'        lErro2 = CF("NFiscalItem_Automatico",lNumIntItemNFiscal)
'        If lErro2 <> SUCESSO Then gError 7 '???

'        objItensNFiscal.lNumIntDoc = lNumIntItemNFiscal
        
'        'Insere itens em ItemNFiscal
'        With objItensNFiscal
'            lErro2 = Comando_Executar(lComando2, "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, Almoxarifado,Ccl, ContaContabil, PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem, ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            .lNumIntNF, .iItem, .iStatus, .sProduto, .sUMVenda, .dQuantidade, .iAlmoxarifado, .sCcl, sContaContabil, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntDocOrig)
'        End With
'        If lErro2 <> AD_SQL_SUCESSO Then gError 8 '???

'        'Verifica se não é uma nota cancelada
'        If objNFiscal.iStatus <> STATUS_CANCELADO Then
        
        
'''        objNFiscal.ColItensNF.Add objItensNFiscal
        
        'Grava a Estatística do Produto Vendido
        lErro2 = CF("ProdutoVendido_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_DOC)
        If lErro2 <> SUCESSO Then gError 9 '???
    
        'end if
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9 '???
    
    Loop
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 11 '???
  
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    lErro = Rotina_Aviso(vbOKOnly, "A criação dos Itens das Notas Fiscais foi realizada com Sucesso!!!")
      
    ConhecFreteFat_Atualiza_Frete2 = SUCESSO

    Exit Function

Erro_ConhecFreteFat_Atualiza_Frete:
    
    ConhecFreteFat_Atualiza_Frete2 = gErr
    
    Select Case gErr
    
        Case 1, 2
            lErro = Rotina_Erro(vbOKOnly, "Erro na Abertura de Comando.", gErr)
            
        Case 3, 4, 9
            lErro = Rotina_Erro(vbOKOnly, "Erro na leitura da tabela de Notas Fiscais", gErr)
        
        Case 5, 6, 7
        
        Case 8
            lErro = Rotina_Erro(vbOKOnly, "Erro na Inserção da tabela de Itens de Notas Fiscais", gErr)

        Case 10
            lErro = Rotina_Erro(vbOKOnly, "Erro na abertura de Transacao.", gErr)
        
       Case 11
            lErro = Rotina_Erro(vbOKOnly, "Erro no Commit da Transacao.", gErr)
        
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select

    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

'*************Funções referentes ao tab de cobrança****************

'Private Function Saida_Celula_Cobrador(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_Cobrador
'
'    Set objGridInt.objControle = Cobrador
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 95210
'
'    Carrega_CarteiraCobrador (Codigo_Extrai(Cobrador))
'    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_CarteiraCobranca_Col) = CarteiraCobrador.List(0)
'
'    Saida_Celula_Cobrador = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Cobrador:
'
'    Saida_Celula_Cobrador = gErr
'
'    Select Case gErr
'
'        Case 95210
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_CarteiraCobrador(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_CarteiraCobrador
'
'    Set objGridInt.objControle = CarteiraCobrador
'
'    'Guarda a opcao selecionada
'    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_CarteiraCobranca_Col) = CarteiraCobrador.Text
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 95211
'
'    Saida_Celula_CarteiraCobrador = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_CarteiraCobrador:
'
'    Saida_Celula_CarteiraCobrador = gErr
'
'    Select Case gErr
'
'        Case 95211
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub Cobrador_Click()
'    iAlterado = REGISTRO_ALTERADO
'    CobrancaAutomatica.Value = vbUnchecked
'End Sub
'Public Sub CarteiraCobrador_Click()
'    iAlterado = REGISTRO_ALTERADO
'    CobrancaAutomatica.Value = vbUnchecked
'End Sub
'
'Public Sub CarteiraCobrador_Change()
'    iAlterado = REGISTRO_ALTERADO
'    CobrancaAutomatica.Value = vbUnchecked
'End Sub
'Public Sub Cobrador_Change()
'    iAlterado = REGISTRO_ALTERADO
'    CobrancaAutomatica.Value = vbUnchecked
'End Sub
'
'Private Function Carrega_Cobradores() As Long
''Carrega a combo de Cobrador
'
'Dim lErro As Long
'Dim colCodigoNome As New AdmColCodigoNome
'Dim objCodigoNome As New AdmCodigoNome
'Dim iIndice As Integer
'
'On Error GoTo Erro_Carrega_Cobradores
'
'    'Leitura dos códigos e descrições dos Bancos BD
'    '?????? Achar constante para nomered de banco com valor = 20
'    lErro = CF("Cod_Nomes_Le", "Cobradores", "Codigo", "NomeReduzido", 20, colCodigoNome)
'    If lErro <> SUCESSO Then Error 95214
'
'   'Preenche ComboBox com código e nome dos Bancos Cobradores
'    For iIndice = 1 To colCodigoNome.Count
'        Set objCodigoNome = colCodigoNome(iIndice)
'        Cobrador.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
'        Cobrador.ItemData(Cobrador.NewIndex) = objCodigoNome.iCodigo
'    Next
'
'    Carrega_Cobradores = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Cobradores:
'
'    Carrega_Cobradores = gErr
'
'    Select Case gErr
'
'        Case 95214 'Tratado na rotina chamada
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Sub Carrega_CarteiraCobrador(iCobrador As Integer)
''Carrega a combo de carteiras de Cobradores
'
'Dim objCobrador As New ClassCobrador
'Dim lErro As Long
'Dim objCarteiraCobrador As New ClassCarteiraCobrador
'Dim objCarteiraCobranca As New ClassCarteiraCobranca
'Dim sListBoxItem As String
'Dim colCarteirasCobrador As New Collection
'
'On Error GoTo Erro_Carrega_CarteiraCobrador
'
'    'Limpa a Combo de Carteiras
'    CarteiraCobrador.Clear
'
'    'Passa o Código do Cobrador que está na tela para o Obj
'    objCobrador.iCodigo = iCobrador
'
'    'Lê os dados do Cobrador
'    lErro = CF("Cobrador_Le", objCobrador)
'    If lErro <> SUCESSO And lErro <> 19294 Then gError 95215
'
'    'Se o Cobrador não estiver cadastrado
'    If lErro = 19294 Then objCobrador.iCodigo = COBRADOR_PROPRIA_EMPRESA
'
'    'Le as carteiras associadas ao Cobrador
'    lErro = CF("Cobrador_Le_Carteiras", objCobrador, colCarteirasCobrador)
'    If lErro <> SUCESSO And lErro <> 23500 Then gError 95218
'
'    If lErro = SUCESSO Then
'
'        'Preencher a Combo
'        For Each objCarteiraCobrador In colCarteirasCobrador
'
'            objCarteiraCobranca.iCodigo = objCarteiraCobrador.iCodCarteiraCobranca
'
'            lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
'            If lErro <> SUCESSO And lErro <> 23413 Then gError 95216
'
'            'Carteira não está cadastrado
'            If lErro = 23413 Then gError 95217
'
'            'Concatena Código e a Descricao da carteira
'            sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
'            sListBoxItem = sListBoxItem & SEPARADOR & objCarteiraCobranca.sDescricao
'
'            CarteiraCobrador.AddItem sListBoxItem
'            CarteiraCobrador.ItemData(CarteiraCobrador.NewIndex) = objCarteiraCobranca.iCodigo
'
'        Next
'
'    End If
'    '??????
'    'Seleciona uma das Carteiras
'    If objCobrador.iCodigo = COBRADOR_PROPRIA_EMPRESA Then Cobrador.Text = Cobrador.List(0)
'    If CarteiraCobrador.ListCount <> 0 Then CarteiraCobrador.ListIndex = 0
'
'    Exit Sub
'
'Erro_Carrega_CarteiraCobrador:
'
'    Select Case gErr
'
'        Case 15836, 95218, 95216
'            Cobrador.SetFocus
'
'        Case 95215
'            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", gErr, Cobrador.Text)
'            Cobrador.SetFocus
'
'        Case 95217
'            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", gErr, objCarteiraCobranca.iCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Cobrador_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Cobrador_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Cobrador_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Cobrador
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub CarteiraCobrador_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub CarteiraCobrador_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub CarteiraCobrador_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = CarteiraCobrador
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'

'Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
''Inicializa o Grid de Parcelas
'
'    'Form do Grid
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add ("Parcela")
'    objGridInt.colColuna.Add ("Vencimento")
'    objGridInt.colColuna.Add ("Valor")
'    objGridInt.colColuna.Add ("Cobrador")
'    objGridInt.colColuna.Add ("Carteira")
'    objGridInt.colColuna.Add ("Desconto 1 Tipo")
'    objGridInt.colColuna.Add ("Desc.1 Data")
'    objGridInt.colColuna.Add ("Desc.1 Valor")
'    objGridInt.colColuna.Add ("Desc.1 %")
'    objGridInt.colColuna.Add ("Desconto 2 Tipo")
'    objGridInt.colColuna.Add ("Desc.2 Data")
'    objGridInt.colColuna.Add ("Desc.2 Valor")
'    objGridInt.colColuna.Add ("Desc.2 %")
'    objGridInt.colColuna.Add ("Desconto 3 Tipo")
'    objGridInt.colColuna.Add ("Desc.3 Data")
'    objGridInt.colColuna.Add ("Desc.3 Valor")
'    objGridInt.colColuna.Add ("Desc.3 %")
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (DataVencimento.Name)
'    objGridInt.colCampo.Add (ValorParcela.Name)
'    objGridInt.colCampo.Add (Cobrador.Name)
'    objGridInt.colCampo.Add (CarteiraCobrador.Name)
'    objGridInt.colCampo.Add (Desconto1Codigo.Name)
'    objGridInt.colCampo.Add (Desconto1Ate.Name)
'    objGridInt.colCampo.Add (Desconto1Valor.Name)
'    objGridInt.colCampo.Add (Desconto1Percentual.Name)
'    objGridInt.colCampo.Add (Desconto2Codigo.Name)
'    objGridInt.colCampo.Add (Desconto2Ate.Name)
'    objGridInt.colCampo.Add (Desconto2Valor.Name)
'    objGridInt.colCampo.Add (Desconto2Percentual.Name)
'    objGridInt.colCampo.Add (Desconto3Codigo.Name)
'    objGridInt.colCampo.Add (Desconto3Ate.Name)
'    objGridInt.colCampo.Add (Desconto3Valor.Name)
'    objGridInt.colCampo.Add (Desconto3Percentual.Name)
'
'    'Colunas do Grid
'    iGrid_Vencimento_col = 1
'    iGrid_ValorParcela_Col = 2
'    iGrid_Cobranca_Col = 3
'    iGrid_CarteiraCobranca_Col = 4
'    iGrid_Desc1Codigo_Col = 5
'    iGrid_Desc1Ate_Col = 6
'    iGrid_Desc1Valor_Col = 7
'    iGrid_Desc1Percentual_Col = 8
'    iGrid_Desc2Codigo_Col = 9
'    iGrid_Desc2Ate_Col = 10
'    iGrid_Desc2Valor_Col = 11
'    iGrid_Desc2Percentual_Col = 12
'    iGrid_Desc3Codigo_Col = 13
'    iGrid_Desc3Ate_Col = 14
'    iGrid_Desc3Valor_Col = 15
'    iGrid_Desc3Percentual_Col = 16
'
'    'Grid do GridInterno
'    objGridInt.objGrid = GridParcelas
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 6
'
'    'Largura da primeira coluna
'    GridParcelas.ColWidth(0) = 700
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'
'    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_Parcelas = SUCESSO
'
'    Exit Function
'
'End Function
'
'
'Function Preenche_DescontoPadrao(iLinha As Integer) As Long
'
'Dim lErro As Long
'Dim colDescontoPadrao As New ColDesconto
'Dim iIndice1 As Integer
'Dim iIndice2 As Integer
'Dim iColuna  As Integer
'Dim dtDataVencimento As Date
'Dim dPercentual As Double
'Dim dValorParcela As Double
'Dim sValorDesconto As String
'
'On Error GoTo Erro_Preenche_DescontoPadrao
'
'    'Se a data de referencia estiver preenchida
'    If Len(Trim(DataReferencia.ClipText)) > 0 Then
'
'        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_col))
'        lErro = CF("Parcela_GeraDescontoPadrao", colDescontoPadrao, dtDataVencimento)
'        If lErro <> SUCESSO Then gError 86065
'
'        If colDescontoPadrao.Count > 0 Then
'
'            'Para cada um dos desontos padrão
'            For iIndice1 = 1 To colDescontoPadrao.Count
'
'                'Seleciona a coluna correspondente ao Desconto
'                If iIndice1 = 1 Then iColuna = iGrid_Desc1Codigo_Col
'                If iIndice1 = 2 Then iColuna = iGrid_Desc2Codigo_Col
'                If iIndice1 = 3 Then iColuna = iGrid_Desc3Codigo_Col
'
'                'Seleciona o tipo de desconto
'                For iIndice2 = 0 To Desconto1Codigo.ListCount - 1
'                    If colDescontoPadrao.Item(iIndice1).iCodigo = Desconto1Codigo.ItemData(iIndice2) Then
'                        GridParcelas.TextMatrix(iLinha, iColuna) = Desconto1Codigo.List(iIndice2)
'                        GridParcelas.TextMatrix(iLinha, iColuna + 1) = Format(colDescontoPadrao.Item(iIndice1).dtData, "dd/mm/yyyy")
'                        GridParcelas.TextMatrix(iLinha, iColuna + 3) = Format(colDescontoPadrao.Item(iIndice1).dValor, "Percent")
'
'                        '*** Inicio colocacao Valor Desconto na tela
'                        dPercentual = colDescontoPadrao.Item(iIndice1).dValor
'                        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))
'
'                        'Coloca Valor do Desconto na tela
'                        If dValorParcela > 0 Then
'                            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
'                            GridParcelas.TextMatrix(iLinha, iColuna + 2) = sValorDesconto
'                        End If
'                        '*** Fim colocacao Valor Desconto na tela
'
'                    End If
'                Next
'            Next
'
'        End If
'
'    End If
'
'    Preenche_DescontoPadrao = SUCESSO
'
'    Exit Function
'
'Erro_Preenche_DescontoPadrao:
'
'    Preenche_DescontoPadrao = gErr
'
'    Select Case gErr
'
'        Case 86065
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Preenche_Grid_Parcelas(objNFiscal As ClassNFiscal) As Long
''Preenche o Grid com as Parcelas da Nota Fiscal
'
'Dim objParcela As New ClassParcelaReceber
'Dim iIndice As Integer
'Dim iIndice2 As Integer
'Dim dValorDesconto As Double
'Dim iContador As Integer
'Dim objCarteiraCobranca As New ClassCarteiraCobranca
'Dim lErro As Long
'
'On Error GoTo Erro_Preenche_Grid_Parcelas
'
'    Call Grid_Limpa(objGridParcelas)
'
'    iIndice = 0
'
'    'PAra cada parcela da coleção de parcelas
'    For Each objParcela In objNFiscal.ColParcelaReceber
'
'        iIndice = iIndice + 1
'        'Preenche o grid com os dados da parcela
'        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col) = Format(objParcela.dtDataVencimento, "dd/mm/yyyy")
'        GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcela.dValor, "Standard")
'
'        'Preenche o Cobrador
'        For iContador = 0 To Cobrador.ListCount - 1
'            If Cobrador.ItemData(iContador) = objParcela.iCobrador Then
'                GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col) = Cobrador.List(iContador)
'                Exit For
'            End If
'        Next
'
'        'Atribui o codigo para o obj a ser passado como parametro
'        objCarteiraCobranca.iCodigo = objParcela.iCarteiraCobranca
'
'        'Le a carteira de cobranca a partir de um código
'        lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
'        If lErro <> SUCESSO And lErro <> 23413 Then gError 95213
'
'        'Preenche a carteira de cobranca
'        GridParcelas.TextMatrix(iIndice, iGrid_CarteiraCobranca_Col) = objCarteiraCobranca.iCodigo & SEPARADOR & objCarteiraCobranca.sDescricao
'
'        If objParcela.dtDesconto1Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col) = Format(objParcela.dtDesconto1Ate, "dd/mm/yyyy")
'        If objParcela.dtDesconto2Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col) = Format(objParcela.dtDesconto2Ate, "dd/mm/yyyy")
'        If objParcela.dtDesconto3Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col) = Format(objParcela.dtDesconto3Ate, "dd/mm/yyyy")
'        If objParcela.iDesconto1Codigo = VALOR_FIXO Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(objParcela.dDesconto1Valor, "Standard")
'        ElseIf objParcela.iDesconto1Codigo = Percentual Or objParcela.iDesconto1Codigo = PERC_ANT_DIA Or objParcela.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col) = Format(objParcela.dDesconto1Valor, "Percent")
'            '*** Inicio código p/ colocar Valor Desconto
'            If objParcela.dValor > 0 Then
'                dValorDesconto = objParcela.dDesconto1Valor * objParcela.dValor
'                GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(dValorDesconto, "Standard")
'            End If
'            '*** Fim
'        End If
'        If objParcela.iDesconto2Codigo = VALOR_FIXO Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(objParcela.dDesconto2Valor, "Standard")
'        ElseIf objParcela.iDesconto2Codigo = Percentual Or objParcela.iDesconto2Codigo = PERC_ANT_DIA Or objParcela.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col) = Format(objParcela.dDesconto2Valor, "Percent")
'            '*** Inicio código p/ colocar Valor Desconto
'            If objParcela.dValor > 0 Then
'                dValorDesconto = objParcela.dDesconto2Valor * objParcela.dValor
'                GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(dValorDesconto, "Standard")
'            End If
'            '*** Fim
'        End If
'        If objParcela.iDesconto3Codigo = VALOR_FIXO Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(objParcela.dDesconto3Valor, "Standard")
'        ElseIf objParcela.iDesconto3Codigo = Percentual Or objParcela.iDesconto3Codigo = PERC_ANT_DIA Or objParcela.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col) = Format(objParcela.dDesconto3Valor, "Percent")
'            '*** Inicio código p/ colocar Valor Desconto
'            If objParcela.dValor > 0 Then
'                dValorDesconto = objParcela.dDesconto3Valor * objParcela.dValor
'                GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(dValorDesconto, "Standard")
'            End If
'            '*** Fim
'        End If
'        For iIndice2 = 0 To Desconto1Codigo.ListCount - 1
'            If Desconto1Codigo.ItemData(iIndice2) = objParcela.iDesconto1Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col) = Desconto1Codigo.List(iIndice2)
'            If Desconto2Codigo.ItemData(iIndice2) = objParcela.iDesconto2Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col) = Desconto2Codigo.List(iIndice2)
'            If Desconto3Codigo.ItemData(iIndice2) = objParcela.iDesconto3Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col) = Desconto3Codigo.List(iIndice2)
'        Next
'
'    Next
'
'    objGridParcelas.iLinhasExistentes = iIndice
'
'    Preenche_Grid_Parcelas = SUCESSO
'
'    Exit Function
'
'Erro_Preenche_Grid_Parcelas:
'
'    Preenche_Grid_Parcelas = gErr
'
'    Select Case gErr
'
'        Case 95213
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_CondicaoPagamento() As Long
''Carrega na combo as Condições de Pagamento existentes
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
'    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then gError 86066
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'        'Adiciona novo ítem na List da Combo CondicaoPagamento
'        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo
'    Next
'
'    Carrega_CondicaoPagamento = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_CondicaoPagamento:
'
'    Carrega_CondicaoPagamento = gErr
'
'    Select Case gErr
'
'        Case 86066
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Function Carrega_TipoDesconto() As Long
''Carrega na combo os Tipos de Desconto existentes
'
'Dim lErro As Long
'Dim objCodDescricao As New AdmCodigoNome
'Dim colCodigoDescricao As AdmColCodigoNome
'
'On Error GoTo Erro_Carrega_TipoDesconto
'
'    Set colCodigoDescricao = gobjCRFAT.colTiposDesconto
'
'    For Each objCodDescricao In colCodigoDescricao
'        'Adiciona o ítem nas List's das Combos de Tipos Desconto
'        Desconto1Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        Desconto1Codigo.ItemData(Desconto1Codigo.NewIndex) = objCodDescricao.iCodigo
'        Desconto2Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        Desconto2Codigo.ItemData(Desconto2Codigo.NewIndex) = objCodDescricao.iCodigo
'        Desconto3Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        Desconto3Codigo.ItemData(Desconto3Codigo.NewIndex) = objCodDescricao.iCodigo
'    Next
'
'    Carrega_TipoDesconto = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_TipoDesconto:
'
'    Carrega_TipoDesconto = gErr
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Valida_Grid_Parcelas() As Long
''Valida os dados do Grid de Parcelas
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim dSomaParcelas As Double
'Dim dValorIRRF As Double
'Dim dValorTotal As Double
'Dim dtDataEmissao As Date
'Dim dtDataVencimento As Date
'Dim iTamanho As Integer
'Dim iTipo As Integer, dValorPagar As Double
'Dim dPercAcrecFin As Double
'Dim iDesconto As Integer
'Dim dtDataDesconto As Date, dValorINSSRetido As Double
'
'On Error GoTo Erro_Valida_Grid_Parcelas
'
'    'Verifica se alguma parcela foi informada
'    If objGridParcelas.iLinhasExistentes = 0 Then gError 86063
'
'    dSomaParcelas = 0
'
'    'Para cada Parcela do grid de parcelas
'    For iIndice = 1 To objGridParcelas.iLinhasExistentes
'
'        dtDataEmissao = StrParaDate(DataEmissao.Text)
'        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))
'
'        'verifica se o vencimento e o valor da parcela estão preenchidos
'        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))) = 0 Then gError 86042
'        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 86043
'        'verifica se a data de vencimento da parcela é menor que a data de emissão
'        If dtDataVencimento < dtDataEmissao Then gError 86044
'        'Se o desconto 1 da parcela está preenchido
'        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
'            iDesconto = 1
'            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
'            'Verifica se a data do desconto está preenchida
'            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 86045
'
'            dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
'            'Valida a data do desconto
'            If dtDataDesconto > dtDataVencimento Then gError 86046
'            'Se o tipo de desconto for por valor
'            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
'                'Verifica se o valor do desconto está preenchido
'                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 86047
'            Else
'                'Verifica se o percentual de desconto está preenchido
'                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col))) = 0 Then gError 86048
'            End If
'            'Se o desconto 2 está preenchido
'            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
'                iDesconto = 2
'                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
'                'Verifica se a data de desconto está preenchida
'                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 86049
'                'Verifica se a data de desconto está ordenada ou se é igual ao desconto anterior
'                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 86050
'                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 86051
'                'Se o desconto for do tipo valor
'                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
'                    'Verifica se o valor está preenchido
'                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 86052
'                Else
'                    'Verifica se o percentual está preenchido
'                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col))) = 0 Then gError 86053
'                End If
'                'Valida a data de desconto
'                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
'                If dtDataDesconto > dtDataVencimento Then gError 86054
'                'Se o desconto 3 está preenchido
'                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
'                    iDesconto = 3
'                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
'                    'Verifica se a data de desconto está preenchida
'                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 86055
'                    'Verifica se a data de desconto está ordenada ou se é igual a do desconto anterior
'                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 86056
'                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 86057
'                    'Se o desconto for do tipo valor
'                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
'                        'Verifica se valor de desconto está preenchido
'                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 86058
'                    Else
'                        'verifica se o percentuial de desconto está preenchido
'                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col))) = 0 Then gError 86059
'                    End If
'                    'Valida a data de desconto
'                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
'                    If dtDataDesconto > dtDataVencimento Then gError 86060
'                End If
'            End If
'        End If
'        'Verifica se as data de vencimentos das parcelas estão ordenadas
'        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_col)) Then gError 86061
'       'Faz a soma do total das parcelas
'        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
'
'    Next
'
'    If INSSRetido.Value = vbChecked Then
'        dValorINSSRetido = StrParaDbl(ValorINSS.Text)
'    End If
'
'    'Valor a Pagar
'    dValorPagar = Round(StrParaDbl(ValorTotal.Caption) - dValorINSSRetido, 2)
'
'    'Verifica se o total das parcelas cobre o valor da nota fiscal
'    If Format(dValorPagar, "Standard") <> Format(dSomaParcelas, "Standard") Then gError 86062
'
'    Valida_Grid_Parcelas = SUCESSO
'
'    Exit Function
'
'Erro_Valida_Grid_Parcelas:
'
'    Valida_Grid_Parcelas = gErr
'
'    Select Case gErr
'
'        Case 86063
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)
'
'        Case 86042
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)
'
'        Case 86044
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)
'
'        Case 86061
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA", gErr)
'
'        Case 86043
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)
'
'        Case 86062
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)
'
'        Case 86045, 86049, 86055
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)
'
'        Case 86047, 86048, 86052, 86053, 86058, 86059
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)
'
'        Case 86050, 86056
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)
'
'        Case 86051, 86057
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)
'
'        Case 86046, 86054, 86060
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_SUPERIOR_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Move_GridParcelas_Memoria(objNFiscal As ClassNFiscal) As Long
''Move as Parcelas do Grid para a Memória
'
'Dim iIndice As Integer
'Dim lTamanho As Long
'Dim objParcela As ClassParcelaReceber
'Dim dtDataReferencia As Date
'Dim dtDataEmissao As Date
'Dim lErro As Long
'
'On Error GoTo Erro_Move_GridParcelas_Memoria
'
'    dtDataReferencia = StrParaDate(DataReferencia.Text)
'    dtDataEmissao = StrParaDate(DataEmissao.Text)
'
'    If dtDataReferencia <> DATA_NULA Then
'        If dtDataReferencia < dtDataEmissao Then gError 86064
'    End If
'
'    'Se não há parcelas a recolher, sai da função
'    If objGridParcelas.iLinhasExistentes <> 0 Then
'
'        'Para cada parcela do grid
'        For iIndice = 1 To objGridParcelas.iLinhasExistentes
'
'            Set objParcela = New ClassParcelaReceber
'
'            objParcela.iNumParcela = iIndice
'
'            'recolhe os dados da parcela
'            objParcela.dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))
'            objParcela.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
'            objParcela.iDesconto1Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
'            objParcela.iDesconto2Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
'            objParcela.iDesconto3Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
'            objParcela.dtDesconto1Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
'            objParcela.dtDesconto2Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
'            objParcela.dtDesconto3Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
'            objParcela.iCobrador = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))
'            objParcela.iCarteiraCobranca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_CarteiraCobranca_Col))
'
'            'Se o tipo de desconto for de Valor recolhe a coluna valor
'            'Senão recolhe a coluna percentual
'            If objParcela.iDesconto1Codigo = VALOR_FIXO Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
'                objParcela.dDesconto1Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))
'            ElseIf objParcela.iDesconto1Codigo = Percentual Or objParcela.iDesconto1Codigo = PERC_ANT_DIA Or objParcela.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
'                lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col)))
'                If lTamanho > 0 Then objParcela.dDesconto1Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col))
'            End If
'
'            'Se o tipo de desconto for de Valor recolhe a coluna valor
'            'Senão recolhe a coluna percentual
'            If objParcela.iDesconto2Codigo = VALOR_FIXO Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
'                objParcela.dDesconto2Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))
'            ElseIf objParcela.iDesconto2Codigo = Percentual Or objParcela.iDesconto2Codigo = PERC_ANT_DIA Or objParcela.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
'                lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col)))
'                If lTamanho > 0 Then objParcela.dDesconto2Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col))
'            End If
'
'            'Se o tipo de desconto for de Valor recolhe a coluna valor
'            'Senão recolhe a coluna percentual
'            If objParcela.iDesconto3Codigo = VALOR_FIXO Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
'                objParcela.dDesconto3Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))
'            ElseIf objParcela.iDesconto3Codigo = Percentual Or objParcela.iDesconto3Codigo = PERC_ANT_DIA Or objParcela.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
'                lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col)))
'                If lTamanho > 0 Then objParcela.dDesconto3Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col))
'            End If
'
'            'Adiciona a parcela na coleção de parcelas da Nota Fiscal
'        'Adiciona a parcela na coleção de parcelas da Nota Fiscal
'        With objParcela
'            objNFiscal.ColParcelaReceber.Add 0, 0, iIndice, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, .iCarteiraCobranca, .iCobrador, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0
'        End With
'
'        Next
'
'    End If
'
'    Move_GridParcelas_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_GridParcelas_Memoria:
'
'    Move_GridParcelas_Memoria = gErr
'
'    Select Case gErr
'
'        Case 86064
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'End Function

'Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
''Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas
'
'Dim lErro As Long
'Dim dValorPagar As Double
'Dim colValorParcelas As New Collection
'Dim colDataVencimento As New Collection
'Dim dtDataReferencia As Date
'Dim iIndice As Integer
'Dim iTamanho As Integer, dValorINSSRetido As Double
'Dim iPosicao As Integer
'Dim objFilialCliente As New ClassFilialCliente
'Dim iContador As Integer
'Dim objCliente As New ClassCliente
'Dim sCliente As String
'
'On Error GoTo Erro_GridParcelas_Preenche
'
'    'Limpa o GridParcelas
'    Call Grid_Limpa(objGridParcelas)
'
'    'Número de Parcelas
'    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas
'
'    If INSSRetido.Value = vbChecked Then
'        dValorINSSRetido = StrParaDbl(ValorINSS.Text)
'    End If
'
'    'Valor a Pagar
'    dValorPagar = Round(StrParaDbl(ValorTotal.Caption) - dValorINSSRetido, 2)
'
'    'Se Valor a Pagar for positivo
'    If dValorPagar > 0 Then
'
'        'Verifica preenchimento de Cliente
'        If Len(Trim(Cliente.ClipText)) > 0 Then
'
'            objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
'            sCliente = Cliente.Text
'
'            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
'            If lErro <> SUCESSO Then gError 95205
'
'        End If
'
'        If objFilialCliente.iCodCobrador = 0 Then objFilialCliente.iCodCobrador = COBRADOR_PROPRIA_EMPRESA
'
'        'Calcula os valores das Parcelas
'        lErro = CF("Parcelas_Calcula", dValorPagar, objCondicaoPagto.iNumeroParcelas, colValorParcelas)
'        If lErro <> SUCESSO Then gError 86023
'
'        'Coloca os valores das Parcelas no Grid Parcelas
'        For iIndice = 1 To objGridParcelas.iLinhasExistentes
'            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(colValorParcelas(iIndice), "Standard")
'
'            For iContador = 0 To Cobrador.ListCount - 1
'                If Cobrador.ItemData(iContador) = objFilialCliente.iCodCobrador Then
'                    GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col) = Cobrador.List(iContador)
'                    Exit For
'                End If
'            Next
'
'            Carrega_CarteiraCobrador (objFilialCliente.iCodCobrador)
'            GridParcelas.TextMatrix(iIndice, iGrid_CarteiraCobranca_Col) = CarteiraCobrador.List(0)
'
'        Next
'
'    End If
'
'    'Se Data Emissão estiver preenchida
'    If Len(Trim(DataReferencia.ClipText)) > 0 Then
'
'        dtDataReferencia = CDate(DataReferencia.Text)
'
'        'Calcula Datas de Vencimento das Parcelas
'        lErro = CF("Parcelas_DatasVencimento", objCondicaoPagto, dtDataReferencia, colDataVencimento)
'        If lErro <> SUCESSO Then gError 86024
'
'        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
'        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
'
'            'Coloca Data de Vencimento no Grid Parcelas
'            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col) = Format(colDataVencimento(iIndice), "dd/mm/yyyy")
'
'        Next
'
'    End If
'
'    For iIndice = 1 To objGridParcelas.iLinhasExistentes
'
'        lErro = Preenche_DescontoPadrao(iIndice)
'        If lErro <> SUCESSO Then gError 86025
'    Next
'    GridParcelas_Preenche = SUCESSO
'
'    Exit Function
'
'Erro_GridParcelas_Preenche:
'
'    GridParcelas_Preenche = gErr
'
'    Select Case gErr
'
'        Case 86023, 86024, 86025
'
'        Case 95205
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'End Function
'
'Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_GridParcelas
'
'    'Verifica qual a coluna atual do Grid
'    Select Case objGridInt.objGrid.Col
'        'Data de Vencimento
'        Case iGrid_Vencimento_col
'            lErro = Saida_Celula_DataVencimento(objGridInt)
'            If lErro <> SUCESSO Then gError 86012
'        'VAlor da Parcela
'        Case iGrid_ValorParcela_Col
'            lErro = Saida_Celula_ValorParcela(objGridInt)
'            If lErro <> SUCESSO Then gError 86013
'        'Descontos da PArcela
'        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
'            lErro = Saida_Celula_TipoDesconto(objGridInt)
'            If lErro <> SUCESSO Then gError 86014
'        'Datas de desconto da Parcela
'        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
'            lErro = Saida_Celula_DescontoData(objGridInt)
'            If lErro <> SUCESSO Then gError 86015
'        'VAlores dos descontos da parcela
'        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
'            lErro = Saida_Celula_DescontoValor(objGridInt)
'            If lErro <> SUCESSO Then gError 86016
'        'Percentuais de desconto da parcela.
'        Case iGrid_Desc1Percentual_Col, iGrid_Desc2Percentual_Col, iGrid_Desc3Percentual_Col
'            lErro = Saida_Celula_DescontoPerc(objGridInt)
'            If lErro <> SUCESSO Then gError 86017
'                'Banco Cobrador
'        Case iGrid_Cobranca_Col
'            lErro = Saida_Celula_Cobrador(objGridInt)
'            If lErro <> SUCESSO Then gError 95208
'        'Carteira de Cobranca
'        Case iGrid_CarteiraCobranca_Col
'            lErro = Saida_Celula_CarteiraCobrador(objGridInt)
'            If lErro <> SUCESSO Then gError 95209
'
'    End Select
'
'    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 86018
'
'    Saida_Celula_GridParcelas = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_GridParcelas:
'
'    Saida_Celula_GridParcelas = gErr
'
'    Select Case gErr
'
'        Case 86012, 86013, 86014, 86017, 86015, 86016, 95208, 95209
'
'        Case 86018
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub CondPagtoLabel_DblClick()
'
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim colSelecao As New Collection
'
'    'Se Condição de Pagto estiver preenchida, extrai o código
'    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
'        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
'    End If
'
'    'Chama a Tela CondicoesPagamentoCRLista
'    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)
'
'End Sub
'
'Private Sub objEventoCondPagto_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objCondicaoPagto As ClassCondicaoPagto
'
'On Error GoTo Erro_objEventoCondPagto_evSelecao
'
'    Set objCondicaoPagto = obj1
'
'    'Preenche campo CondicaoPagamento
'    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
'
'    'Se Valor nao estiver preenchido
'    If Len(Trim(ValorTotal.Caption)) = 0 Then Exit Sub
'
'    'Se DataEmissao estiver preenchida e Valor for positivo
'    If Len(Trim(DataEmissao.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then
'
'        'Preenche GridParcelas a partir da Condição de Pagto
'        lErro = Cobranca_Automatica()
'        If lErro <> SUCESSO Then gError 42138
'
'    End If
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoCondPagto_evSelecao:
'
'    Select Case gErr
'
'        Case 42138
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'     End Select
'
'     Exit Sub
'
'End Sub
'
'
'Public Sub CondicaoPagamento_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub CondicaoPagamento_Click()
'
'Dim lErro As Long
'Dim objCondicaoPagto As New ClassCondicaoPagto
'
'On Error GoTo Erro_CondicaoPagamento_Click
'
'    'Verifica se alguma Condição foi selecionada
'    If CondicaoPagamento.ListIndex = -1 Then Exit Sub
'
'    'Passa o código da Condição para objCondicaoPagto
'    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)
'
'    'Lê Condição de Pagamento à partir do código
'    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
'    If lErro <> SUCESSO And lErro <> 19205 Then gError 42247
'
'    'Não encontrou a Condição de Pagamento --> erro
'    If lErro = 19205 Then gError 42248
'
'    'Testa se ValorTotal está preenchido
'    If Len(Trim(ValorTotal)) > 0 Then
'
'        'Preenche o GridParcelas
'        lErro = Cobranca_Automatica()
'        If lErro <> SUCESSO Then gError 42249
'
'    End If
'
'    iAlterado = REGISTRO_ALTERADO
'
'    Exit Sub
'
'Erro_CondicaoPagamento_Click:
'
'    Select Case gErr
'
'        Case 42247, 42249
'
'        Case 42248
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'      End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub CondicaoPagamento_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_Condicaopagamento_Validate
'
'    'Verifica se a Condicaopagamento foi preenchida
'    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub
'
'    'Verifica se é uma Condicaopagamento selecionada
'    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub
'
'    'Tenta selecionar na combo
'    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42250
'
'    'Se não encontra valor que contém CÓDIGO, mas extrai o código
'    If lErro = 6730 Then
'
'        objCondicaoPagto.iCodigo = iCodigo
'
'        'Lê Condição Pagamento no BD
'        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
'        If lErro <> SUCESSO And lErro <> 19205 Then gError 42251
'
'        'Não encontrou a Condição de Pagamento
'        If lErro = 19205 Then gError 42252
'
'        'Testa se pode ser usada em Contas a Receber
'        If objCondicaoPagto.iEmRecebimento = 0 Then gError 42253
'
'        'Coloca na Tela
'        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida
'
'        'Se ValorTotal e DataEmissao estiverem preenchidos, preenche GridParcelas
'        If Len(Trim(ValorTotal)) > 0 Then
'            If Len(Trim(DataReferencia.ClipText)) > 0 And CDbl(ValorTotal.Caption) > 0 Then
'
'                'Preenche o GridParcelas
'                lErro = Cobranca_Automatica()
'                If lErro <> SUCESSO Then gError 42254
'
'            End If
'        End If
'
'    End If
'
'    'Não encontrou o valor que era STRING
'    If lErro = 6731 Then gError 42255
'
'    Exit Sub
'
'Erro_Condicaopagamento_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'       Case 42250, 42251, 42254
'
'       Case 42252
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
'            End If
'
'        Case 42253
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)
'
'        Case 42255
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function Cobranca_Automatica() As Long
''recalcula o tab de cobranca
'
'Dim lErro As Long
'Dim objCondicaoPagto As New ClassCondicaoPagto
'
'On Error GoTo Erro_Cobranca_Automatica
'
'    'Se automática estiver selecionada e a condicao de pagamento estiver preenchida
'    If CobrancaAutomatica.Value = 1 And Len(Trim(CondicaoPagamento.Text)) <> 0 Then
'        'Pega a condicao de pagamento da tela
'        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
'        'Lê a condição de pagamento
'        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
'        If lErro <> SUCESSO And lErro <> 19205 Then gError 46187
'        If lErro <> SUCESSO Then gError 46188
'        'Preenche o grid parcelas de acordo com a condição de pagamento
'        lErro = GridParcelas_Preenche(objCondicaoPagto)
'        If lErro <> SUCESSO Then gError 46189
'        CobrancaAutomatica.Tag = GRID_CHECKBOX_ATIVO
'        CobrancaAutomatica.Value = 1
'
'    End If
'
'    Cobranca_Automatica = SUCESSO
'
'    Exit Function
'
'Erro_Cobranca_Automatica:
'
'    Cobranca_Automatica = gErr
'
'    Select Case gErr
'
'        Case 46187, 46189
'
'        Case 46188
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub CobrancaAutomatica_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) > 0 And CobrancaAutomatica.Tag <> GRID_CHECKBOX_ATIVO Then
'        Call Cobranca_Automatica
'    Else
'        CobrancaAutomatica.Tag = ""
'    End If
'
'End Sub
'
'Public Sub DataReferencia_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giDataReferenciaAlterada = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DataReferencia_GotFocus()
'
'Dim iDataAux As Integer
'
'    iDataAux = giDataReferenciaAlterada
'    Call MaskEdBox_TrataGotFocus(DataReferencia, iAlterado)
'    giDataReferenciaAlterada = iDataAux
'
'End Sub
'
'Public Sub DataReferencia_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dtDataEmissao As Date
'Dim dtDataReferencia As Date
'Dim objCondicaoPagto As New ClassCondicaoPagto
'
'On Error GoTo Erro_DataReferencia_Validate
'
'    If giDataReferenciaAlterada <> REGISTRO_ALTERADO Then Exit Sub
'
'    If Len(Trim(DataReferencia.ClipText)) > 0 Then
'
'        'Critica a data digitada
'        lErro = Data_Critica(DataReferencia.Text)
'        If lErro <> SUCESSO Then gError 26713
'
'        'Compara com data de emissão
'        If Len(Trim(DataEmissao.ClipText)) > 0 Then
'
'            dtDataEmissao = CDate(DataEmissao.Text)
'            dtDataReferencia = CDate(DataReferencia.Text)
'
'            If dtDataEmissao > dtDataReferencia Then gError 26714
'
'        End If
'
'
'    End If
'
'    giDataReferenciaAlterada = 0
'
'    'Preenche o GridParcelas
'    lErro = Cobranca_Automatica()
'    If lErro <> SUCESSO Then gError 25436
'
'    Exit Sub
'
'Erro_DataReferencia_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 25436 'Tratado na rotina chamada
'
'        Case 26713
'
'        Case 26714
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub BotaoDataReferenciaUp_Click()
'
'Dim lErro As Long
'Dim sData As String
'Dim bCancel As Boolean
'
'On Error GoTo Erro_BotaoDataReferenciaUp_Click
'
'    sData = DataReferencia.Text
'
'    'aumenta a data em um dia
'    lErro = Data_Up_Down_Click(DataReferencia, AUMENTA_DATA)
'    If lErro <> SUCESSO Then gError 26716
'
'    Call DataReferencia_Validate(bCancel)
'
'    If bCancel = True Then DataReferencia.Text = sData
'
'    Exit Sub
'
'Erro_BotaoDataReferenciaUp_Click:
'
'    Select Case gErr
'
'        Case 26716
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub BotaoDataReferenciaDown_Click()
'
'Dim lErro As Long
'Dim bCancel As Boolean
'Dim sData As String
'
'On Error GoTo Erro_BotaoDataReferenciaDown_Click
'
'    sData = DataReferencia.Text
'
'    'diminui a data em um dia
'    lErro = Data_Up_Down_Click(DataReferencia, DIMINUI_DATA)
'    If lErro <> SUCESSO Then gError 26715
'
'    Call DataReferencia_Validate(bCancel)
'
'    If bCancel = True Then DataReferencia.Text = sData
'
'    Exit Sub
'
'Erro_BotaoDataReferenciaDown_Click:
'
'    Select Case gErr
'
'        Case 26715
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub GridParcelas_Click()
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridParcelas_GotFocus()
'
'    Call Grid_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_EnterCell()
'
'    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
'
'End Sub
'
'Public Sub GridParcelas_LeaveCell()
'
'    Call Saida_Celula(objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridParcelas_Validate(Cancel As Boolean)
'
'    Call Grid_Libera_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_RowColChange()
'
'    Call Grid_RowColChange(objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_Scroll()
'
'    Call Grid_Scroll(objGridParcelas)
'
'End Sub
'
'Public Sub DataVencimento_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DataVencimento_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub DataVencimento_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub DataVencimento_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = DataVencimento
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub ValorParcela_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorParcela_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub ValorParcela_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub ValorParcela_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = ValorParcela
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto1Ate_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto1Ate_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Ate_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Ate_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto1Ate
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto1Codigo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Desconto1Codigo_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Codigo_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Codigo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto1Codigo
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto1Percentual_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto1Percentual_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Percentual_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto1Percentual
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto1Valor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto1Valor_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Valor_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Valor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto1Valor
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto2Ate_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto2Ate_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Ate_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Ate_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto2Ate
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto2Codigo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Desconto2Codigo_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Codigo_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Codigo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto2Codigo
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto2Percentual_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto2Percentual_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Percentual_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto2Percentual
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto2Valor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto2Valor_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Valor_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Valor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto2Valor
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto3Ate_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto3Ate_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Ate_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Ate_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto3Ate
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto3Codigo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto3Codigo_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Codigo_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Codigo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto3Codigo
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto3Percentual_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto3Percentual_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Percentual_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto3Percentual
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto3Valor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto3Valor_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Valor_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Valor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto3Valor
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Private Function Saida_Celula_ValorParcela(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Valor da Parcela que está deixando de sser a corrente
'
'Dim lErro As Long
'Dim dColunaSoma As Double
'Dim iIndice As Integer
'Dim iColDescPerc As Integer
'Dim iColTipoDesconto As Integer
'Dim lTamanho As Long
'Dim dPercentual As Double
'Dim dValorParcela As Double
'Dim sValorDesconto As String
'Dim iTipoDesconto As Integer
'
'On Error GoTo Erro_Saida_Celula_ValorParcela
'
'    Set objGridInt.objControle = ValorParcela
'
'    'Verifica se valor está preenchido
'    If Len(ValorParcela.ClipText) > 0 Then
'
'        'Critica se valor é positivo
'        lErro = Valor_Positivo_Critica(ValorParcela.Text)
'        If lErro <> SUCESSO Then gError 86026
'
'        ValorParcela.Text = Format(ValorParcela.Text, "Standard")
'
'        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then
'
'           ' CobrancaAutomatica.Value = vbUnchecked
'
'            '***Código para colocar valores de desconto
'            dValorParcela = StrParaDbl(ValorParcela.Text)
'            If dValorParcela > 0 Then
'
'                'Vai varrer todos os 3 descontos para colocar valores
'                For iIndice = 1 To 3
'
'                    Select Case iIndice
'                        Case 1
'                            iColDescPerc = iGrid_Desc1Percentual_Col
'                            iColTipoDesconto = iGrid_Desc1Codigo_Col
'                        Case 2
'                            iColDescPerc = iGrid_Desc2Percentual_Col
'                            iColTipoDesconto = iGrid_Desc2Codigo_Col
'                        Case 3
'                            iColDescPerc = iGrid_Desc3Percentual_Col
'                            iColTipoDesconto = iGrid_Desc3Codigo_Col
'                    End Select
'
'                    iTipoDesconto = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iColTipoDesconto))
'                    lTamanho = Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc)))
'
'                    'Coloca valor de desconto na tela
'                    If (iTipoDesconto = Percentual Or iTipoDesconto = PERC_ANT_DIA Or iTipoDesconto = PERC_ANT_DIA_UTIL) And lTamanho > 0 Then
'                        dPercentual = PercentParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc))
'                        sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
'                        GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc - 1) = sValorDesconto
'                    End If
'
'                Next
'
'            End If
'            '***Fim Código para colocar valores de desconto
'
'        End If
'
'        'Acrescenta uma linha no Grid se for o caso
'        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
'            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
'            'Coloca desconto padrao (le em CPRConfig)
'            lErro = Preenche_DescontoPadrao(GridParcelas.Row)
'            If lErro <> SUCESSO Then gError 86027
'
'        End If
'
'    Else
'        '***Código para colocar valores de desconto
'        'Limpa Valores de Desconto
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc1Valor_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc2Valor_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc3Valor_Col) = ""
'        '***Fim Código para colocar valores de desconto
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 86028
'
'    Saida_Celula_ValorParcela = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_ValorParcela:
'
'    Saida_Celula_ValorParcela = gErr
'
'    Select Case gErr
'
'        Case 86026, 86028, 86027
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function

'******************************************
'4 eventos do controle do Grid de Comissoes: DiretoIndireto
'******************************************

'Alterado por Tulio em 03/04

Private Sub DiretoIndireto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiretoIndireto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub DiretoIndireto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub DiretoIndireto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = DiretoIndireto
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeMarca_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca, "AVISO_CRIAR_VOLUMEMARCA")
    If lErro <> SUCESSO Then gError 102417
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102417
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154844)

    End Select

End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeEspecie_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie, "AVISO_CRIAR_VOLUMEESPECIE")
    If lErro <> SUCESSO Then gError 102416
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102416
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154845)

    End Select

End Sub

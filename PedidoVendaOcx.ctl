VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PedidoVendaOcx 
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Base ICMS Subst"
      Height          =   4965
      Index           =   3
      Left            =   30
      TabIndex        =   62
      Top             =   1110
      Visible         =   0   'False
      Width           =   9390
      Begin VB.Frame Frame6 
         Caption         =   "Complemento"
         Height          =   2220
         Index           =   7
         Left            =   60
         TabIndex        =   192
         Top             =   2685
         Width           =   9285
         Begin VB.ComboBox CanalVenda 
            Height          =   315
            Left            =   1500
            TabIndex        =   207
            Top             =   1755
            Width           =   1440
         End
         Begin VB.TextBox Mensagem 
            Height          =   885
            Left            =   1500
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   193
            Top             =   270
            Width           =   7725
         End
         Begin MSMask.MaskEdBox PedidoCliente 
            Height          =   300
            Left            =   4305
            TabIndex        =   194
            Top             =   1800
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   300
            Left            =   4305
            TabIndex        =   195
            Top             =   1335
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cubagem 
            Height          =   300
            Left            =   7785
            TabIndex        =   196
            Top             =   1365
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PedidoRepr 
            Height          =   300
            Left            =   7785
            TabIndex        =   205
            Top             =   1800
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   300
            Left            =   1500
            TabIndex        =   208
            Top             =   1335
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ped. Representante:"
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
            Index           =   35
            Left            =   6000
            TabIndex        =   206
            Top             =   1845
            Width           =   1770
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
            Height          =   195
            Index           =   56
            Left            =   465
            TabIndex        =   202
            Top             =   1380
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Peso Líquido:"
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
            Index           =   55
            Left            =   3105
            TabIndex        =   201
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label CanalVendaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Canal de Venda:"
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
            TabIndex        =   200
            Top             =   1830
            Width           =   1425
         End
         Begin VB.Label MensagemLabel 
            AutoSize        =   -1  'True
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
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   199
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Cliente:"
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
            Left            =   3000
            TabIndex        =   198
            Top             =   1845
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cubagem:"
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
            Index           =   57
            Left            =   6930
            TabIndex        =   197
            Top             =   1410
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados de Entrega"
         Height          =   2055
         Left            =   45
         TabIndex        =   122
         Top             =   -30
         Width           =   9300
         Begin VB.Frame Frame1 
            Caption         =   "Redespacho"
            Height          =   1065
            Index           =   13
            Left            =   5460
            TabIndex        =   183
            Top             =   765
            Width           =   3780
            Begin VB.ComboBox TranspRedespacho 
               Height          =   315
               Left            =   1515
               TabIndex        =   185
               Top             =   285
               Width           =   2220
            End
            Begin VB.CheckBox RedespachoCli 
               Caption         =   "por conta do cliente"
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
               Left            =   225
               TabIndex        =   184
               Top             =   705
               Width           =   2100
            End
            Begin VB.Label TranspRedLabel 
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
               Left            =   90
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   186
               Top             =   330
               Width           =   1365
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frete por conta do"
            Height          =   870
            Index           =   1
            Left            =   60
            TabIndex        =   125
            Top             =   600
            Width           =   2745
            Begin VB.ComboBox TipoFrete 
               Height          =   315
               Left            =   45
               Style           =   2  'Dropdown List
               TabIndex        =   227
               Top             =   390
               Width           =   2670
            End
         End
         Begin VB.ComboBox FilialEntrega 
            Height          =   315
            Left            =   1830
            TabIndex        =   63
            Top             =   225
            Width           =   3630
         End
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   6975
            TabIndex        =   64
            Top             =   270
            Width           =   2235
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   4140
            MaxLength       =   10
            TabIndex        =   65
            Top             =   690
            Width           =   1290
         End
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   4140
            TabIndex        =   66
            Top             =   1125
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDownEntregaPV 
            Height          =   300
            Left            =   2940
            TabIndex        =   187
            TabStop         =   0   'False
            Top             =   1605
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntregaPV 
            Height          =   300
            Left            =   1815
            TabIndex        =   188
            Top             =   1605
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data de Entrega:"
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
            Index           =   5
            Left            =   300
            TabIndex        =   189
            Top             =   1635
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial para Entrega:"
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
            Index           =   100
            Left            =   150
            TabIndex        =   130
            Top             =   285
            Width           =   1620
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
            Left            =   5550
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   131
            Top             =   315
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Placa Veículo:"
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
            Index           =   54
            Left            =   2820
            TabIndex        =   132
            Top             =   750
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "U.F. da Placa:"
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
            Index           =   53
            Left            =   2835
            TabIndex        =   133
            Top             =   1185
            Width           =   1245
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Volumes"
         Height          =   630
         Left            =   45
         TabIndex        =   120
         Top             =   2040
         Width           =   9300
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   3120
            TabIndex        =   68
            Top             =   210
            Width           =   1650
         End
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   5535
            TabIndex        =   69
            Top             =   210
            Width           =   1650
         End
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7755
            MaxLength       =   20
            TabIndex        =   70
            Top             =   210
            Width           =   1440
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1395
            TabIndex        =   67
            Top             =   210
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            Index           =   32
            Left            =   7365
            TabIndex        =   134
            Top             =   270
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
            Height          =   195
            Index           =   50
            Left            =   300
            TabIndex        =   135
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label Label1 
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
            Index           =   51
            Left            =   2295
            TabIndex        =   136
            Top             =   270
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
            Height          =   195
            Index           =   52
            Left            =   4935
            TabIndex        =   137
            Top             =   270
            Width           =   600
         End
      End
   End
   Begin VB.CommandButton BotaoAnaliseVenda 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6015
      TabIndex        =   279
      Top             =   180
      Width           =   330
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4965
      Index           =   2
      Left            =   45
      TabIndex        =   36
      Top             =   1095
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton BotaoImportarItens 
         Caption         =   "Importar"
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
         Left            =   8430
         TabIndex        =   61
         Top             =   4605
         Width           =   870
      End
      Begin VB.CommandButton BotaoInfoAdicItem 
         Caption         =   "Inf. Adicionais Item"
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
         Left            =   6600
         TabIndex        =   60
         Top             =   4605
         Width           =   1800
      End
      Begin VB.Frame Frame2 
         Caption         =   "Totais"
         Height          =   1290
         Index           =   4
         Left            =   45
         TabIndex        =   238
         Top             =   3225
         Width           =   9285
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   90
            TabIndex        =   51
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   75
            TabIndex        =   239
            Top             =   405
            Visible         =   0   'False
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   2745
            TabIndex        =   53
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   1410
            TabIndex        =   52
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercDescontoItens 
            Height          =   285
            Left            =   4065
            TabIndex        =   54
            ToolTipText     =   "Percentual de desconto dos itens"
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDescontoItens 
            Height          =   285
            Left            =   5400
            TabIndex        =   55
            ToolTipText     =   "Soma dos descontos dos itens"
            Top             =   915
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   263
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   262
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   90
            TabIndex        =   260
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1410
            TabIndex        =   259
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2745
            TabIndex        =   258
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4065
            TabIndex        =   257
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   256
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label Label1 
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
            Height          =   180
            Index           =   28
            Left            =   165
            TabIndex        =   255
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   27
            Left            =   1470
            TabIndex        =   254
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "BC ICMS ST"
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
            Index           =   26
            Left            =   2745
            TabIndex        =   253
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ICMS ST"
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
            Index           =   25
            Left            =   4080
            TabIndex        =   252
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   24
            Left            =   8100
            TabIndex        =   251
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ISS"
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
            Index           =   22
            Left            =   6735
            TabIndex        =   250
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "% Desconto"
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
            Index           =   14
            Left            =   4125
            TabIndex        =   249
            Top             =   705
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   12
            Left            =   5430
            TabIndex        =   248
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label ISSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5400
            TabIndex        =   247
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Base ISS"
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
            Index           =   15
            Left            =   5430
            TabIndex        =   246
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   11
            Left            =   105
            TabIndex        =   245
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   10
            Left            =   1470
            TabIndex        =   244
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   7
            Left            =   2790
            TabIndex        =   243
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   4
            Left            =   6735
            TabIndex        =   242
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Height          =   180
            Index           =   21
            Left            =   8085
            TabIndex        =   241
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label ValorProdutos2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   240
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   261
            Top             =   405
            Width           =   1140
         End
      End
      Begin VB.CommandButton BotaoEntrega 
         Caption         =   "Datas de Entrega"
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
         Left            =   3315
         TabIndex        =   214
         Top             =   4605
         Width           =   1650
      End
      Begin VB.CommandButton BotaoKitVenda 
         Caption         =   "Kits de Venda"
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
         Left            =   945
         TabIndex        =   57
         Top             =   4605
         Width           =   1365
      End
      Begin VB.CommandButton BotaoGrade 
         Caption         =   "Grade"
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
         Left            =   0
         TabIndex        =   56
         Top             =   4605
         Width           =   885
      End
      Begin VB.CommandButton BotaoEstoqueProd 
         Caption         =   "Estoque-Produto"
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
         Left            =   4995
         TabIndex        =   59
         Top             =   4605
         Width           =   1575
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
         Height          =   345
         Left            =   2370
         TabIndex        =   58
         Top             =   4605
         Width           =   915
      End
      Begin VB.Frame Frame2 
         Caption         =   "Itens"
         Height          =   3270
         Index           =   3
         Left            =   45
         TabIndex        =   37
         Top             =   -45
         Width           =   9285
         Begin MSMask.MaskEdBox ComissaoItemPV 
            Height          =   255
            Left            =   7440
            TabIndex        =   278
            Top             =   1890
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   450
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
         Begin VB.ComboBox TabPrecoItemPV 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "PedidoVendaOcx.ctx":0000
            Left            =   2400
            List            =   "PedidoVendaOcx.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   277
            Top             =   1800
            Width           =   1800
         End
         Begin MSMask.MaskEdBox PrecoTotalB 
            Height          =   225
            Left            =   7500
            TabIndex        =   264
            Top             =   1650
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSMask.MaskEdBox Prioridade 
            Height          =   255
            Left            =   3885
            TabIndex        =   224
            Top             =   1290
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantFatAMais 
            Height          =   225
            Left            =   4350
            TabIndex        =   222
            Top             =   1860
            Width           =   1500
            _ExtentX        =   2646
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
         Begin MSMask.MaskEdBox PercentMenosReceb 
            Height          =   255
            Left            =   6000
            TabIndex        =   221
            Top             =   1845
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   450
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
         Begin VB.ComboBox RecebForaFaixa 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "PedidoVendaOcx.ctx":0004
            Left            =   5055
            List            =   "PedidoVendaOcx.ctx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   219
            Top             =   1335
            Width           =   2235
         End
         Begin MSMask.MaskEdBox PercentMaisReceb 
            Height          =   255
            Left            =   5820
            TabIndex        =   220
            Top             =   1230
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   450
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
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3990
            MaxLength       =   250
            TabIndex        =   48
            Top             =   660
            Width           =   2145
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "PedidoVendaOcx.ctx":0043
            Left            =   1575
            List            =   "PedidoVendaOcx.ctx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   240
            Width           =   720
         End
         Begin MSMask.MaskEdBox QuantCancelada 
            Height          =   225
            Left            =   7140
            TabIndex        =   44
            Top             =   360
            Width           =   1485
            _ExtentX        =   2619
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
         Begin MSMask.MaskEdBox QuantFaturada 
            Height          =   225
            Left            =   6990
            TabIndex        =   50
            Top             =   720
            Width           =   1500
            _ExtentX        =   2646
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
         Begin MSMask.MaskEdBox QuantReservadaPV 
            Height          =   225
            Left            =   5400
            TabIndex        =   49
            Top             =   720
            Width           =   1500
            _ExtentX        =   2646
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
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   2640
            TabIndex        =   47
            Top             =   660
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
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   1410
            TabIndex        =   46
            Top             =   630
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
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   255
            TabIndex        =   45
            Top             =   660
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   4185
            TabIndex        =   42
            Top             =   375
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2580
            TabIndex        =   41
            Top             =   300
            Width           =   1500
            _ExtentX        =   2646
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   315
            TabIndex        =   39
            Top             =   330
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   5685
            TabIndex        =   43
            Top             =   360
            Width           =   1185
            _ExtentX        =   2090
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
            Height          =   2175
            Left            =   45
            TabIndex        =   38
            Top             =   195
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4605
      Index           =   5
      Left            =   45
      TabIndex        =   90
      Top             =   1185
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   4125
         Index           =   0
         Left            =   60
         TabIndex        =   126
         Top             =   390
         Width           =   9060
         Begin VB.ComboBox DiretoIndireto 
            Height          =   315
            ItemData        =   "PedidoVendaOcx.ctx":0047
            Left            =   5070
            List            =   "PedidoVendaOcx.ctx":0051
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   915
            Width           =   1335
         End
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
            Height          =   645
            Left            =   7440
            Picture         =   "PedidoVendaOcx.ctx":0067
            Style           =   1  'Graphical
            TabIndex        =   174
            Top             =   3300
            Width           =   1500
         End
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
            Left            =   120
            TabIndex        =   169
            Top             =   3195
            Width           =   6975
            Begin VB.Label TotalValorBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1440
               TabIndex        =   176
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total:"
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
               Left            =   360
               TabIndex        =   175
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label TotalValorComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5640
               TabIndex        =   173
               Top             =   345
               Width           =   1155
            End
            Begin VB.Label TotalPercentualComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3960
               TabIndex        =   172
               Top             =   345
               Width           =   825
            End
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
               Left            =   5040
               TabIndex        =   171
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
               Left            =   2880
               TabIndex        =   170
               Top             =   360
               Width           =   1095
            End
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   180
            Left            =   435
            TabIndex        =   92
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualBaixa 
            Height          =   180
            Left            =   7020
            TabIndex        =   93
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox ValorBaixa 
            Height          =   180
            Left            =   7875
            TabIndex        =   94
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   180
            Left            =   3825
            TabIndex        =   178
            Top             =   165
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   180
            Left            =   2700
            TabIndex        =   179
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   180
            Left            =   1815
            TabIndex        =   180
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   180
            Left            =   5880
            TabIndex        =   181
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   180
            Left            =   5025
            TabIndex        =   182
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
            Height          =   1950
            Left            =   150
            TabIndex        =   95
            Top             =   360
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   3440
            _Version        =   393216
            Rows            =   11
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CheckBox ComissaoAutomatica 
         Caption         =   "Calcula comissão automaticamente"
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   525
         TabIndex        =   91
         Top             =   135
         Value           =   1  'Checked
         Width           =   3360
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4650
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   1170
      Width           =   9195
      Begin VB.CommandButton BotaoTodosPedidos 
         Caption         =   "Todos os Pedidos de Venda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7305
         TabIndex        =   203
         Top             =   3990
         Width           =   1740
      End
      Begin VB.CheckBox FaturaIntegral 
         Caption         =   "Só libera o pedido integralmente"
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
         Left            =   240
         TabIndex        =   34
         Top             =   4140
         Width           =   3300
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   900
         Index           =   6
         Left            =   210
         TabIndex        =   22
         Top             =   2010
         Width           =   8865
         Begin VB.CommandButton BotaoContato 
            Caption         =   "Clientes Futuros"
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
            Left            =   7395
            TabIndex        =   209
            Top             =   210
            Width           =   1290
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5175
            TabIndex        =   26
            Top             =   345
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1815
            TabIndex        =   24
            Top             =   360
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
            Index           =   13
            Left            =   4650
            TabIndex        =   25
            Top             =   405
            Width           =   465
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1110
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   23
            Top             =   405
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Preços"
         Height          =   900
         Index           =   2
         Left            =   210
         TabIndex        =   27
         Top             =   2970
         Width           =   8865
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   4485
            TabIndex        =   31
            Top             =   345
            Width           =   1815
         End
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   1320
            TabIndex        =   29
            Top             =   345
            Width           =   1875
         End
         Begin MSMask.MaskEdBox PercAcrescFin 
            Height          =   315
            Left            =   7995
            TabIndex        =   33
            Top             =   345
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
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
            Format          =   "#0.#0\%"
            PromptChar      =   " "
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
            Left            =   3390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   405
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Acrésc Financ:"
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
            Left            =   6495
            TabIndex        =   32
            Top             =   405
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tabela Preço:"
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
            Index           =   62
            Left            =   90
            TabIndex        =   28
            Top             =   405
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1770
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   150
         Width           =   8865
         Begin VB.Frame FrameCodBase 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   330
            Left            =   6615
            TabIndex        =   274
            Top             =   270
            Visible         =   0   'False
            Width           =   2100
            Begin VB.Label CodBase 
               BorderStyle     =   1  'Fixed Single
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
               Left            =   1005
               TabIndex        =   276
               Top             =   0
               Width           =   1080
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Cód.Base:"
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
               TabIndex        =   275
               Top             =   90
               Width           =   870
            End
         End
         Begin VB.Frame FrameParc 
            BorderStyle     =   0  'None
            Caption         =   "FrameParc"
            Height          =   480
            Left            =   3300
            TabIndex        =   272
            Top             =   195
            Visible         =   0   'False
            Width           =   1095
            Begin MSMask.MaskEdBox Parc 
               Height          =   300
               Left            =   690
               TabIndex        =   12
               Top             =   120
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   1
               Mask            =   "#"
               PromptChar      =   " "
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Parc.:"
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
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   273
               Top             =   180
               Width           =   525
            End
         End
         Begin VB.CommandButton BotaoExportar 
            Caption         =   "Exportar"
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
            Left            =   7485
            TabIndex        =   228
            Top             =   675
            Visible         =   0   'False
            Width           =   1215
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
            Left            =   3765
            TabIndex        =   21
            Top             =   1230
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   5190
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1215
            Width           =   2550
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2910
            Picture         =   "PedidoVendaOcx.ctx":0611
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Numeração Automática"
            Top             =   330
            Width           =   300
         End
         Begin VB.ComboBox FilialFaturamento 
            Height          =   315
            ItemData        =   "PedidoVendaOcx.ctx":06FB
            Left            =   5190
            List            =   "PedidoVendaOcx.ctx":06FD
            TabIndex        =   18
            Top             =   765
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1815
            TabIndex        =   11
            Top             =   315
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   6225
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   5175
            TabIndex        =   15
            Top             =   315
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   1815
            TabIndex        =   19
            Top             =   1245
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LblNatOpInternaEspelho 
            AutoSize        =   -1  'True
            Caption         =   "Natureza de Oper.:"
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
            TabIndex        =   226
            Top             =   810
            Width           =   1650
         End
         Begin VB.Label NatOpInternaEspelho 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1815
            TabIndex        =   225
            Top             =   750
            Width           =   525
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
            Left            =   1080
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   211
            Top             =   1290
            Width           =   675
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
            Index           =   41
            Left            =   4560
            TabIndex        =   210
            Top             =   1275
            Width           =   570
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   4365
            TabIndex        =   14
            Top             =   360
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
            Left            =   1035
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   10
            Top             =   345
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial Faturamento:"
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
            Index           =   17
            Left            =   3570
            TabIndex        =   17
            Top             =   795
            Width           =   1575
         End
      End
      Begin MSMask.MaskEdBox PrioridadePadrao 
         Height          =   315
         Left            =   5325
         TabIndex        =   35
         Top             =   4110
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Prioridade Padrão:"
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
         Index           =   40
         Left            =   3645
         TabIndex        =   223
         Top             =   4170
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4815
      Index           =   7
      Left            =   30
      TabIndex        =   105
      Top             =   1260
      Visible         =   0   'False
      Width           =   9240
      Begin VB.CommandButton BotaoImprimirConf 
         Height          =   345
         Left            =   5460
         Picture         =   "PedidoVendaOcx.ctx":06FF
         Style           =   1  'Graphical
         TabIndex        =   271
         ToolTipText     =   "Imprimir relatório para conferência do estoque"
         Top             =   4320
         Width           =   420
      End
      Begin VB.CommandButton BotaoRefazAlocacao 
         Caption         =   "Refaz Reservas do Pedido"
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
         Left            =   6075
         TabIndex        =   204
         Top             =   4365
         Width           =   2955
      End
      Begin VB.CommandButton BotaoLibera 
         Caption         =   "Libera Reservas do Pedido"
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
         Left            =   6090
         TabIndex        =   116
         Top             =   3945
         Width           =   2940
      End
      Begin VB.CommandButton BotaoReserva 
         Caption         =   "Reserva dos Produtos Pedidos"
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
         Left            =   6105
         TabIndex        =   115
         Top             =   3540
         Width           =   2940
      End
      Begin VB.Frame Frame7 
         Caption         =   "Reserva dos Produtos"
         Height          =   3450
         Left            =   105
         TabIndex        =   121
         Top             =   45
         Width           =   8940
         Begin VB.TextBox Responsavel 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5790
            TabIndex        =   113
            Top             =   795
            Width           =   2115
         End
         Begin MSMask.MaskEdBox UnidadeMedEst 
            Height          =   225
            Left            =   7965
            TabIndex        =   112
            Top             =   360
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
         Begin MSMask.MaskEdBox DataValidade 
            Height          =   225
            Left            =   6675
            TabIndex        =   111
            Top             =   345
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
         Begin MSMask.MaskEdBox ProdutoAlmox 
            Height          =   225
            Left            =   1335
            TabIndex        =   107
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almox 
            Height          =   225
            Left            =   2550
            TabIndex        =   108
            Top             =   390
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantReservar 
            Height          =   225
            Left            =   3825
            TabIndex        =   109
            Top             =   390
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox QuantReservada 
            Height          =   225
            Left            =   5250
            TabIndex        =   110
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
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
         Begin MSMask.MaskEdBox ItemPedido 
            Height          =   225
            Left            =   765
            TabIndex        =   106
            Top             =   390
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
         Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
            Height          =   2805
            Left            =   180
            TabIndex        =   114
            Top             =   285
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4948
            _Version        =   393216
            Rows            =   11
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   735
         Left            =   105
         TabIndex        =   123
         Top             =   3540
         Width           =   5790
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   34
            Left            =   240
            TabIndex        =   138
            Top             =   330
            Width           =   735
         End
         Begin VB.Label ProdutoDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1140
            TabIndex        =   139
            Top             =   300
            Width           =   4395
         End
      End
   End
   Begin VB.CheckBox ImprimirConfGravacao 
      Caption         =   "Imprimir relat. p/conf. do estoque ao gravar"
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
      Left            =   5265
      TabIndex        =   230
      Top             =   6150
      Width           =   4170
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4980
      Index           =   4
      Left            =   90
      TabIndex        =   71
      Top             =   1095
      Visible         =   0   'False
      Width           =   9270
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
         Left            =   15
         TabIndex        =   72
         Top             =   -15
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   4755
         Left            =   15
         TabIndex        =   127
         Top             =   180
         Width           =   9270
         Begin VB.CommandButton BotaoDataRefFluxoUp 
            Height          =   150
            Left            =   8970
            Picture         =   "PedidoVendaOcx.ctx":0801
            Style           =   1  'Graphical
            TabIndex        =   216
            TabStop         =   0   'False
            Top             =   675
            Width           =   240
         End
         Begin VB.CommandButton BotaoDataRefFluxoDown 
            Height          =   150
            Left            =   8970
            Picture         =   "PedidoVendaOcx.ctx":085B
            Style           =   1  'Graphical
            TabIndex        =   215
            TabStop         =   0   'False
            Top             =   825
            Width           =   240
         End
         Begin VB.CommandButton BotaoTipoPagto 
            Caption         =   "Detalhamento Tipo de Pagto (F5)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7365
            TabIndex        =   213
            Top             =   4275
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.ComboBox TipoPagto 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "PedidoVendaOcx.ctx":08B5
            Left            =   3045
            List            =   "PedidoVendaOcx.ctx":08C5
            TabIndex        =   212
            Top             =   2310
            Width           =   1965
         End
         Begin VB.CommandButton BotaoDataReferenciaDown 
            Height          =   150
            Left            =   3105
            Picture         =   "PedidoVendaOcx.ctx":0902
            Style           =   1  'Graphical
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   855
            Width           =   240
         End
         Begin VB.CommandButton BotaoDataReferenciaUp 
            Height          =   150
            Left            =   3105
            Picture         =   "PedidoVendaOcx.ctx":095C
            Style           =   1  'Graphical
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   705
            Width           =   240
         End
         Begin VB.ComboBox TipoDesconto1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   77
            Top             =   1215
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3120
            TabIndex        =   78
            Top             =   1515
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   79
            Top             =   1845
            Width           =   1965
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7470
            TabIndex        =   86
            Top             =   1260
            Width           =   900
            _ExtentX        =   1588
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
            Left            =   6105
            TabIndex        =   85
            Top             =   1905
            Width           =   1275
            _ExtentX        =   2249
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
            Left            =   4995
            TabIndex        =   82
            Top             =   1905
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
            TabIndex        =   84
            Top             =   1590
            Width           =   1275
            _ExtentX        =   2249
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
            Left            =   4995
            TabIndex        =   81
            Top             =   1590
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
            Left            =   6120
            TabIndex        =   83
            Top             =   1260
            Width           =   1275
            _ExtentX        =   2249
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
            Left            =   4980
            TabIndex        =   80
            Top             =   1260
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
            Left            =   570
            TabIndex        =   75
            Top             =   1230
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
            Height          =   240
            Left            =   1695
            TabIndex        =   76
            Top             =   1245
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Percentual 
            Height          =   225
            Left            =   7500
            TabIndex        =   87
            Top             =   1605
            Width           =   900
            _ExtentX        =   1588
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
            Left            =   7455
            TabIndex        =   88
            Top             =   1905
            Width           =   900
            _ExtentX        =   1588
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
         Begin MSMask.MaskEdBox DataReferencia 
            Height          =   300
            Left            =   2010
            TabIndex        =   74
            Top             =   705
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2745
            Left            =   45
            TabIndex        =   89
            Top             =   1215
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   4842
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox DataRefFluxo 
            Height          =   300
            Left            =   7875
            TabIndex        =   217
            Top             =   675
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDescontoTit 
            Height          =   300
            Left            =   4470
            TabIndex        =   73
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total a Receber:"
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
            Left            =   6330
            TabIndex        =   237
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label ValorTit 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7875
            TabIndex        =   236
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Original:"
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
            Left            =   705
            TabIndex        =   235
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label ValorOriginalTit 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2010
            TabIndex        =   234
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desconto:"
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
            Index           =   8
            Left            =   3540
            TabIndex        =   233
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data de Referência p/fluxo:"
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
            Index           =   37
            Left            =   5385
            TabIndex        =   218
            Top             =   735
            Width           =   2400
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
            Index           =   9
            Left            =   180
            TabIndex        =   140
            Top             =   765
            Width           =   1740
         End
      End
   End
   Begin VB.CommandButton BotaoInfoAdic 
      Caption         =   "Informações Adicionais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4695
      TabIndex        =   119
      Top             =   75
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4680
      Index           =   8
      Left            =   30
      TabIndex        =   117
      Top             =   1185
      Visible         =   0   'False
      Width           =   9270
      Begin TelasFAT.TabTributacaoFat TabTrib 
         Height          =   4575
         Left            =   90
         TabIndex        =   232
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   8070
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   6450
      ScaleHeight     =   465
      ScaleWidth      =   2940
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   90
      Width           =   3000
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2430
         Picture         =   "PedidoVendaOcx.ctx":09B6
         Style           =   1  'Graphical
         TabIndex        =   270
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1950
         Picture         =   "PedidoVendaOcx.ctx":0B34
         Style           =   1  'Graphical
         TabIndex        =   269
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1485
         Picture         =   "PedidoVendaOcx.ctx":1066
         Style           =   1  'Graphical
         TabIndex        =   268
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   1005
         Picture         =   "PedidoVendaOcx.ctx":11F0
         Style           =   1  'Graphical
         TabIndex        =   267
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   510
         Picture         =   "PedidoVendaOcx.ctx":134A
         Style           =   1  'Graphical
         TabIndex        =   266
         ToolTipText     =   "Imprimir relatório de pedido de venda"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoEmail 
         Height          =   345
         Left            =   45
         Picture         =   "PedidoVendaOcx.ctx":144C
         Style           =   1  'Graphical
         TabIndex        =   265
         ToolTipText     =   "Enviar email"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.CheckBox ImprimeGravacao 
      Caption         =   "Imprimir o pedido ao gravar"
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
      Left            =   105
      TabIndex        =   231
      Top             =   6150
      Width           =   2715
   End
   Begin VB.CheckBox EmailGravacao 
      Caption         =   "Enviar email ao gravar"
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
      Left            =   2895
      TabIndex        =   229
      Top             =   6150
      Width           =   2280
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Private Sub Observacao_Change()Private Sub Observacao_Change()"
      Height          =   4605
      Index           =   6
      Left            =   30
      TabIndex        =   96
      Top             =   1245
      Visible         =   0   'False
      Width           =   9240
      Begin VB.CommandButton BotaoLiberaBloqueio 
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
         Left            =   6825
         Picture         =   "PedidoVendaOcx.ctx":1DEE
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   3840
         Width           =   1650
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Bloqueios"
         Height          =   3630
         Left            =   75
         TabIndex        =   124
         Top             =   90
         Width           =   9120
         Begin VB.TextBox SeqBloqueio 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   1710
            MaxLength       =   250
            TabIndex        =   191
            Top             =   2430
            Width           =   675
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3075
            MaxLength       =   250
            TabIndex        =   190
            Top             =   855
            Width           =   4245
         End
         Begin VB.ComboBox TipoBloqueio 
            Height          =   315
            ItemData        =   "PedidoVendaOcx.ctx":43E8
            Left            =   555
            List            =   "PedidoVendaOcx.ctx":43EA
            TabIndex        =   97
            Top             =   330
            Width           =   1605
         End
         Begin MSMask.MaskEdBox ResponsavelLib 
            Height          =   270
            Left            =   7380
            TabIndex        =   102
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataLiberacao 
            Height          =   270
            Left            =   6150
            TabIndex        =   101
            Top             =   345
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodUsuario 
            Height          =   270
            Left            =   3330
            TabIndex        =   99
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ResponsavelBL 
            Height          =   270
            Left            =   4785
            TabIndex        =   100
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataBloqueio 
            Height          =   270
            Left            =   2190
            TabIndex        =   98
            Top             =   345
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridBloqueio 
            Height          =   2715
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   4789
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
   Begin VB.Frame FrameOrcVenda 
      Caption         =   "Orçamento de Venda"
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4650
      Begin VB.CommandButton BotaoOrcamento 
         Height          =   360
         Left            =   3795
         Picture         =   "PedidoVendaOcx.ctx":43EC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Trazer Dados"
         Top             =   210
         Width           =   360
      End
      Begin VB.CommandButton BotaoVerOrcamento 
         Height          =   360
         Left            =   4185
         Picture         =   "PedidoVendaOcx.ctx":47BE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Trazer a tela de cadastro"
         Top             =   210
         Width           =   360
      End
      Begin VB.ComboBox FilialOrcamento 
         Height          =   315
         Left            =   2295
         TabIndex        =   4
         Top             =   240
         Width           =   1500
      End
      Begin MSMask.MaskEdBox Orcamento 
         Height          =   300
         Left            =   870
         TabIndex        =   2
         Top             =   240
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelFilialOrcamento 
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
         Left            =   1830
         TabIndex        =   3
         Top             =   300
         Width           =   465
      End
      Begin VB.Label OrcamentoLabel 
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
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   1
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   9
      Left            =   90
      TabIndex        =   141
      Top             =   1170
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CheckBox CalculoAuto 
         Caption         =   "Calcula embalagens automaticamente"
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
         Left            =   300
         TabIndex        =   154
         Top             =   30
         Value           =   1  'Checked
         Width           =   3750
      End
      Begin VB.CommandButton BotaoEmbalagens 
         Caption         =   "Produto  X  Embalagens"
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
         Left            =   6360
         TabIndex        =   150
         Top             =   4170
         Width           =   2355
      End
      Begin VB.Frame Frame11 
         Caption         =   "Embalagens"
         Height          =   3705
         Index           =   1
         Left            =   0
         TabIndex        =   142
         Top             =   390
         Width           =   9090
         Begin VB.Frame Frame14 
            Caption         =   "Detalhes do Elemento Selecionado"
            Height          =   705
            Left            =   120
            TabIndex        =   162
            Top             =   2880
            Width           =   6495
            Begin VB.Label DescProduto 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1035
               TabIndex        =   168
               Top             =   300
               Width           =   1935
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   33
               Left            =   135
               TabIndex        =   167
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "UM:"
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
               Index           =   16
               Left            =   3195
               TabIndex        =   166
               Top             =   300
               Width           =   390
            End
            Begin VB.Label UMProduto 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3705
               TabIndex        =   165
               Top             =   300
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Quant:"
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
               Index           =   23
               Left            =   4905
               TabIndex        =   164
               Top             =   300
               Width           =   585
            End
            Begin VB.Label ProdutoQuant 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5565
               TabIndex        =   163
               Top             =   300
               Width           =   765
            End
         End
         Begin MSMask.MaskEdBox Capacidade 
            Height          =   225
            Left            =   3660
            TabIndex        =   143
            Top             =   120
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoEmb 
            Height          =   225
            Left            =   1380
            TabIndex        =   144
            Top             =   150
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UMEmb 
            Height          =   225
            Left            =   2580
            TabIndex        =   145
            Top             =   150
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoLiq 
            Height          =   225
            Left            =   5880
            TabIndex        =   146
            Top             =   150
            Width           =   1050
            _ExtentX        =   1852
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
         Begin MSMask.MaskEdBox ItemEmb 
            Height          =   225
            Left            =   30
            TabIndex        =   147
            Top             =   150
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantEmb 
            Height          =   225
            Left            =   4650
            TabIndex        =   149
            Top             =   120
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBrutoEmb 
            Height          =   225
            Left            =   7110
            TabIndex        =   151
            Top             =   150
            Width           =   1170
            _ExtentX        =   2064
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
         Begin MSMask.MaskEdBox QuantProduto 
            Height          =   225
            Left            =   600
            TabIndex        =   152
            Top             =   540
            Width           =   1200
            _ExtentX        =   2117
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
         Begin MSMask.MaskEdBox Embalagem 
            Height          =   225
            Left            =   540
            TabIndex        =   153
            Top             =   210
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridEmb 
            Height          =   1980
            Left            =   360
            TabIndex        =   148
            Top             =   270
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   3493
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label PesoBrutoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7470
            TabIndex        =   161
            Top             =   2430
            Width           =   1425
         End
         Begin VB.Label PesoLiqTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4725
            TabIndex        =   160
            Top             =   2430
            Width           =   1425
         End
         Begin VB.Label QuantEmbTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1995
            TabIndex        =   159
            Top             =   2430
            Width           =   1425
         End
         Begin VB.Label EmbTotais 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   158
            Top             =   2490
            Width           =   600
         End
         Begin VB.Label EmbTotais 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Emb.:"
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
            Left            =   960
            TabIndex        =   157
            Top             =   2490
            Width           =   960
         End
         Begin VB.Label EmbTotais 
            AutoSize        =   -1  'True
            Caption         =   "Peso Líq.:"
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
            Left            =   3720
            TabIndex        =   156
            Top             =   2460
            Width           =   900
         End
         Begin VB.Label EmbTotais 
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
            Height          =   195
            Index           =   3
            Left            =   6390
            TabIndex        =   155
            Top             =   2460
            Width           =   1005
         End
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5415
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   9551
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
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
            Caption         =   "Cobrança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bloqueio"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Almoxarifado"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Embalagens"
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
Attribute VB_Name = "PedidoVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTPedidoVenda
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoAnaliseVenda_Click()
    Call objCT.BotaoAnaliseVenda_Click
End Sub

Private Sub BotaoDataReferenciaDown_Click()
     Call objCT.BotaoDataReferenciaDown_Click
End Sub

Private Sub BotaoDataReferenciaUp_Click()
     Call objCT.BotaoDataReferenciaUp_Click
End Sub
'Janaina
Private Sub BotaoEmbalagens_Click()
    Call objCT.BotaoEmbalagens_Click
End Sub

Private Sub BotaoExportar_Click()
    Call objCT.BotaoExportar_Click
End Sub

Private Sub BotaoGrade_Click()
    Call objCT.BotaoGrade_Click
End Sub

Private Sub BotaoImportarItens_Click()
    Call objCT.BotaoImportarItens_Click
End Sub

Private Sub BotaoInfoAdicItem_Click()
    Call objCT.BotaoInfoAdicItem_Click
End Sub

Private Sub BotaoOrcamento_Click()
    Call objCT.BotaoOrcamento_Click
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoVerOrcamento_Click()
    Call objCT.BotaoVerOrcamento_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEntregaPV_Change()
    Call objCT.DataEntregaPV_Change
End Sub

Private Sub DataEntregaPV_GotFocus()
    Call objCT.DataEntregaPV_GotFocus
End Sub

Private Sub DataEntregaPV_Validate(Cancel As Boolean)
    Call objCT.DataEntregaPV_Validate(Cancel)
End Sub

Private Sub DataReferencia_GotFocus()
     Call objCT.DataReferencia_GotFocus
End Sub

Private Sub DiretoIndireto_Change()
    Call objCT.DiretoIndireto_Change
End Sub

Private Sub DiretoIndireto_GotFocus()
    Call objCT.DiretoIndireto_GotFocus
End Sub

Private Sub DiretoIndireto_KeyPress(KeyAscii As Integer)
    Call objCT.DiretoIndireto_KeyPress(KeyAscii)
End Sub

Private Sub DiretoIndireto_Validate(Cancel As Boolean)
    Call objCT.DiretoIndireto_Validate(Cancel)
End Sub

Private Sub FilialFaturamento_Change()
     Call objCT.FilialFaturamento_Change
End Sub

Private Sub FilialFaturamento_Click()
     Call objCT.FilialFaturamento_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub FilialOrcamento_Change()
    Call objCT.FilialOrcamento_Change
End Sub

Private Sub FilialOrcamento_Click()
    Call objCT.FilialOrcamento_Click
End Sub

Private Sub FilialOrcamento_Validate(Cancel As Boolean)
    Call objCT.FilialOrcamento_Validate(Cancel)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridItens_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridItens_Scroll()
     Call objCT.GridItens_Scroll
End Sub

Private Sub LabelCliente_Click()
     Call objCT.LabelCliente_Click
End Sub

Private Sub LabelFilialOrcamento_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelFilialOrcamento, Source, X, Y)
End Sub

Private Sub LabelFilialOrcamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelFilialOrcamento, Button, Shift, X, Y)
End Sub

Private Sub MensagemLabel_Click()
     Call objCT.MensagemLabel_Click
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Private Sub CondPagtoLabel_Click()
     Call objCT.CondPagtoLabel_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub BotaoEstoqueProd_Click()
     Call objCT.BotaoEstoqueProd_Click
End Sub

Function Trata_Parametros(Optional objPedidoVenda As ClassPedidoDeVenda) As Long
     Trata_Parametros = objCT.Trata_Parametros(objPedidoVenda)
End Function

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub Orcamento_Change()
    Call objCT.Orcamento_Change 'Por leo em 26/03/02
End Sub

Private Sub Orcamento_GotFocus()
    Call objCT.Orcamento_GotFocus 'Por leo em 26/03/02
End Sub


Private Sub OrcamentoLabel_Click()
    Call objCT.OrcamentoLabel_Click
End Sub

Private Sub OrcamentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(OrcamentoLabel, Source, X, Y)
End Sub

Private Sub OrcamentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(OrcamentoLabel, Button, Shift, X, Y)
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

Private Sub PlacaUF_Validate(Cancel As Boolean)
     Call objCT.PlacaUF_Validate(Cancel)
End Sub

Private Sub RedespachoCli_Click()
    Call objCT.RedespachoCli_Click
End Sub

Private Sub TipoFrete_Click()
    Call objCT.TipoFrete_Click
End Sub

Private Sub UnidadeMed_Click()
     Call objCT.UnidadeMed_Click
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub TabelaPreco_Click()
     Call objCT.TabelaPreco_Click
End Sub

Private Sub CanalVenda_Change()
     Call objCT.CanalVenda_Change
End Sub

Private Sub CanalVenda_Click()
     Call objCT.CanalVenda_Click
End Sub

Private Sub CobrancaAutomatica_Click()
     Call objCT.CobrancaAutomatica_Click
End Sub

Private Sub CondicaoPagamento_Change()
     Call objCT.CondicaoPagamento_Change
End Sub

Private Sub CondicaoPagamento_Click()
     Call objCT.CondicaoPagamento_Click
End Sub

Private Sub DataReferencia_Change()
     Call objCT.DataReferencia_Change
End Sub

Private Sub FilialEntrega_Change()
     Call objCT.FilialEntrega_Change
End Sub

Private Sub FilialEntrega_Click()
     Call objCT.FilialEntrega_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
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

Private Sub Mensagem_Change()
     Call objCT.Mensagem_Change
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub PedidoCliente_Change()
     Call objCT.PedidoCliente_Change
End Sub

Private Sub PercAcrescFin_Change()
     Call objCT.PercAcrescFin_Change
End Sub

Private Sub ComissaoAutomatica_Click()
     Call objCT.ComissaoAutomatica_Click
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub Transportadora_Click()
     Call objCT.Transportadora_Click
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

Private Sub UnidadeMed_Change()
     Call objCT.UnidadeMed_Change
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

Private Sub Quantidade_Change()
     Call objCT.Quantidade_Change
End Sub

Private Sub Quantidade_GotFocus()
     Call objCT.Quantidade_GotFocus
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
     Call objCT.Quantidade_KeyPress(KeyAscii)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub PrecoUnitario_Change()
     Call objCT.PrecoUnitario_Change
End Sub

Private Sub PrecoUnitario_GotFocus()
     Call objCT.PrecoUnitario_GotFocus
End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoUnitario_KeyPress(KeyAscii)
End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)
     Call objCT.PrecoUnitario_Validate(Cancel)
End Sub

Private Sub PrecoTotal_Change()
     Call objCT.PrecoTotal_Change
End Sub

Private Sub PrecoTotal_GotFocus()
     Call objCT.PrecoTotal_GotFocus
End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoTotal_KeyPress(KeyAscii)
End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)
     Call objCT.PrecoTotal_Validate(Cancel)
End Sub

Private Sub QuantCancelada_Change()
     Call objCT.QuantCancelada_Change
End Sub

Private Sub QuantCancelada_GotFocus()
     Call objCT.QuantCancelada_GotFocus
End Sub

Private Sub QuantCancelada_KeyPress(KeyAscii As Integer)
     Call objCT.QuantCancelada_KeyPress(KeyAscii)
End Sub

Private Sub QuantCancelada_Validate(Cancel As Boolean)
     Call objCT.QuantCancelada_Validate(Cancel)
End Sub

Private Sub QuantReservadaPV_Change()
     Call objCT.QuantReservadaPV_Change
End Sub

Private Sub QuantReservadaPV_GotFocus()
     Call objCT.QuantReservadaPV_GotFocus
End Sub

Private Sub QuantReservadaPV_KeyPress(KeyAscii As Integer)
     Call objCT.QuantReservadaPV_KeyPress(KeyAscii)
End Sub

Private Sub QuantReservadaPV_Validate(Cancel As Boolean)
     Call objCT.QuantReservadaPV_Validate(Cancel)
End Sub

Private Sub QuantFaturada_Change()
     Call objCT.QuantFaturada_Change
End Sub

Private Sub QuantFaturada_GotFocus()
     Call objCT.QuantFaturada_GotFocus
End Sub

Private Sub QuantFaturada_KeyPress(KeyAscii As Integer)
     Call objCT.QuantFaturada_KeyPress(KeyAscii)
End Sub

Private Sub QuantFaturada_Validate(Cancel As Boolean)
     Call objCT.QuantFaturada_Validate(Cancel)
End Sub

Private Sub PercentDesc_Change()
     Call objCT.PercentDesc_Change
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

Private Sub Desconto_Change()
     Call objCT.Desconto_Change
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

Private Sub DataEntrega_Change()
     Call objCT.DataEntrega_Change
End Sub

Private Sub DataEntrega_GotFocus()
     Call objCT.DataEntrega_GotFocus
End Sub

Private Sub DataEntrega_KeyPress(KeyAscii As Integer)
     Call objCT.DataEntrega_KeyPress(KeyAscii)
End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)
     Call objCT.DataEntrega_Validate(Cancel)
End Sub

Private Sub DescricaoProduto_Change()
     Call objCT.DescricaoProduto_Change
End Sub

Private Sub DescricaoProduto_GotFocus()
     Call objCT.DescricaoProduto_GotFocus
End Sub

Private Sub DescricaoProduto_KeyPress(KeyAscii As Integer)
     Call objCT.DescricaoProduto_KeyPress(KeyAscii)
End Sub

Private Sub DescricaoProduto_Validate(Cancel As Boolean)
     Call objCT.DescricaoProduto_Validate(Cancel)
End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)
     Call objCT.TabelaPreco_Validate(Cancel)
End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)
     Call objCT.CondicaoPagamento_Validate(Cancel)
End Sub

Private Sub UpDownEntregaPV_DownClick()
    Call objCT.UpDownEntregaPV_DownClick
End Sub

Private Sub UpDownEntregaPV_UpClick()
    Call objCT.UpDownEntregaPV_UpClick
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTPedidoVenda
    Set objCT.objUserControl = Me
End Sub

Public Property Set objCTTela(ByVal vData As Object)
    Set objCT = vData
End Property

Public Property Get objCTTela() As Object
    Set objCTTela = objCT
End Property

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub ValorDespesas_Change()
     Call objCT.ValorDespesas_Change
End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)
     Call objCT.ValorDespesas_Validate(Cancel)
End Sub

Private Sub ValorFrete_Change()
     Call objCT.ValorFrete_Change
End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)
     Call objCT.ValorFrete_Validate(Cancel)
End Sub

Private Sub ValorProdutos_Change()
     Call objCT.ValorProdutos_Change
End Sub

Private Sub ValorSeguro_Change()
     Call objCT.ValorSeguro_Change
End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)
     Call objCT.ValorSeguro_Validate(Cancel)
End Sub

Private Sub FilialEntrega_Validate(Cancel As Boolean)
     Call objCT.FilialEntrega_Validate(Cancel)
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub CanalVenda_Validate(Cancel As Boolean)
     Call objCT.CanalVenda_Validate(Cancel)
End Sub

Private Sub DataReferencia_Validate(Cancel As Boolean)
     Call objCT.DataReferencia_Validate(Cancel)
End Sub

Private Sub BotaoVendedores_Click()
     Call objCT.BotaoVendedores_Click
End Sub

Private Sub BotaoLiberaBloqueio_Click()
     Call objCT.BotaoLiberaBloqueio_Click
End Sub

Private Sub BotaoLibera_Click()
     Call objCT.BotaoLibera_Click
End Sub

Private Sub BotaoReserva_Click()
     Call objCT.BotaoReserva_Click
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

Private Sub TipoDesconto1_Change()
     Call objCT.TipoDesconto1_Change
End Sub

Private Sub TipoDesconto1_GotFocus()
     Call objCT.TipoDesconto1_GotFocus
End Sub

Private Sub TipoDesconto1_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto1_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto1_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto1_Validate(Cancel)
End Sub

Private Sub TipoDesconto2_Change()
     Call objCT.TipoDesconto2_Change
End Sub

Private Sub TipoDesconto2_GotFocus()
     Call objCT.TipoDesconto2_GotFocus
End Sub

Private Sub TipoDesconto2_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto2_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto2_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto2_Validate(Cancel)
End Sub

Private Sub TipoDesconto3_Change()
     Call objCT.TipoDesconto3_Change
End Sub

Private Sub TipoDesconto3_GotFocus()
     Call objCT.TipoDesconto3_GotFocus
End Sub

Private Sub TipoDesconto3_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto3_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto3_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto3_Validate(Cancel)
End Sub

Private Sub Desconto1Ate_Change()
     Call objCT.Desconto1Ate_Change
End Sub

Private Sub Desconto1Ate_GotFocus()
     Call objCT.Desconto1Ate_GotFocus
End Sub

Private Sub Desconto1Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto1Ate_Validate(Cancel)
End Sub

Private Sub Desconto2Ate_Change()
     Call objCT.Desconto2Ate_Change
End Sub

Private Sub Desconto2Ate_GotFocus()
     Call objCT.Desconto2Ate_GotFocus
End Sub

Private Sub Desconto2Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto2Ate_Validate(Cancel)
End Sub

Private Sub Desconto3Ate_Change()
     Call objCT.Desconto3Ate_Change
End Sub

Private Sub Desconto3Ate_GotFocus()
     Call objCT.Desconto3Ate_GotFocus
End Sub

Private Sub Desconto3Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto3Ate_Validate(Cancel)
End Sub

Private Sub Desconto1Valor_Change()
     Call objCT.Desconto1Valor_Change
End Sub

Private Sub Desconto1Valor_GotFocus()
     Call objCT.Desconto1Valor_GotFocus
End Sub

Private Sub Desconto1Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto1Valor_Validate(Cancel)
End Sub

Private Sub Desconto2Valor_Change()
     Call objCT.Desconto2Valor_Change
End Sub

Private Sub Desconto2Valor_GotFocus()
     Call objCT.Desconto2Valor_GotFocus
End Sub

Private Sub Desconto2Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto2Valor_Validate(Cancel)
End Sub

Private Sub Desconto3Valor_Change()
     Call objCT.Desconto3Valor_Change
End Sub

Private Sub Desconto3Valor_GotFocus()
     Call objCT.Desconto3Valor_GotFocus
End Sub

Private Sub Desconto3Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto3Valor_Validate(Cancel)
End Sub

Private Sub Desconto1Percentual_Change()
     Call objCT.Desconto1Percentual_Change
End Sub

Private Sub Desconto1Percentual_GotFocus()
     Call objCT.Desconto1Percentual_GotFocus
End Sub

Private Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto1Percentual_Validate(Cancel)
End Sub

Private Sub Desconto2Percentual_Change()
     Call objCT.Desconto2Percentual_Change
End Sub

Private Sub Desconto2Percentual_GotFocus()
     Call objCT.Desconto2Percentual_GotFocus
End Sub

Private Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto2Percentual_Validate(Cancel)
End Sub

Private Sub Desconto3Percentual_Change()
     Call objCT.Desconto3Percentual_Change
End Sub

Private Sub Desconto3Percentual_GotFocus()
     Call objCT.Desconto3Percentual_GotFocus
End Sub

Private Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto3Percentual_Validate(Cancel)
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

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_GotFocus()
     Call objCT.Vendedor_GotFocus
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
     Call objCT.Vendedor_KeyPress(KeyAscii)
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Private Sub PercentualComissao_Change()
     Call objCT.PercentualComissao_Change
End Sub

Private Sub PercentualComissao_GotFocus()
     Call objCT.PercentualComissao_GotFocus
End Sub

Private Sub PercentualComissao_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualComissao_KeyPress(KeyAscii)
End Sub

Private Sub PercentualComissao_Validate(Cancel As Boolean)
     Call objCT.PercentualComissao_Validate(Cancel)
End Sub

Private Sub ValorBase_Change()
     Call objCT.ValorBase_Change
End Sub

Private Sub ValorBase_GotFocus()
     Call objCT.ValorBase_GotFocus
End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBase_KeyPress(KeyAscii)
End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)
     Call objCT.ValorBase_Validate(Cancel)
End Sub

Private Sub ValorComissao_Change()
     Call objCT.ValorComissao_Change
End Sub

Private Sub ValorComissao_GotFocus()
     Call objCT.ValorComissao_GotFocus
End Sub

Private Sub ValorComissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorComissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorComissao_Validate(Cancel As Boolean)
     Call objCT.ValorComissao_Validate(Cancel)
End Sub

Private Sub PercentualEmissao_Change()
     Call objCT.PercentualEmissao_Change
End Sub

Private Sub PercentualEmissao_GotFocus()
     Call objCT.PercentualEmissao_GotFocus
End Sub

Private Sub PercentualEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualEmissao_KeyPress(KeyAscii)
End Sub

Private Sub PercentualEmissao_Validate(Cancel As Boolean)
     Call objCT.PercentualEmissao_Validate(Cancel)
End Sub

Private Sub ValorEmissao_Change()
     Call objCT.ValorEmissao_Change
End Sub

Private Sub ValorEmissao_GotFocus()
     Call objCT.ValorEmissao_GotFocus
End Sub

Private Sub ValorEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorEmissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorEmissao_Validate(Cancel As Boolean)
     Call objCT.ValorEmissao_Validate(Cancel)
End Sub

Private Sub PercentualBaixa_Change()
     Call objCT.PercentualBaixa_Change
End Sub

Private Sub PercentualBaixa_GotFocus()
     Call objCT.PercentualBaixa_GotFocus
End Sub

Private Sub PercentualBaixa_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualBaixa_KeyPress(KeyAscii)
End Sub

Private Sub PercentualBaixa_Validate(Cancel As Boolean)
     Call objCT.PercentualBaixa_Validate(Cancel)
End Sub

Private Sub ValorBaixa_Change()
     Call objCT.ValorBaixa_Change
End Sub

Private Sub ValorBaixa_GotFocus()
     Call objCT.ValorBaixa_GotFocus
End Sub

Private Sub ValorBaixa_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBaixa_KeyPress(KeyAscii)
End Sub

Private Sub ValorBaixa_Validate(Cancel As Boolean)
     Call objCT.ValorBaixa_Validate(Cancel)
End Sub

Private Sub GridComissoes_Click()
     Call objCT.GridComissoes_Click
End Sub

Private Sub GridComissoes_GotFocus()
     Call objCT.GridComissoes_GotFocus
End Sub

Private Sub GridComissoes_EnterCell()
     Call objCT.GridComissoes_EnterCell
End Sub

Private Sub GridComissoes_LeaveCell()
     Call objCT.GridComissoes_LeaveCell
End Sub

Private Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridComissoes_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridComissoes_KeyPress(KeyAscii As Integer)
     Call objCT.GridComissoes_KeyPress(KeyAscii)
End Sub

Private Sub GridComissoes_Validate(Cancel As Boolean)
     Call objCT.GridComissoes_Validate(Cancel)
End Sub

Private Sub GridComissoes_RowColChange()
     Call objCT.GridComissoes_RowColChange
End Sub

Private Sub GridComissoes_Scroll()
     Call objCT.GridComissoes_Scroll
End Sub

Private Sub TipoBloqueio_Change()
     Call objCT.TipoBloqueio_Change
End Sub

Private Sub TipoBloqueio_GotFocus()
     Call objCT.TipoBloqueio_GotFocus
End Sub

Private Sub TipoBloqueio_KeyPress(KeyAscii As Integer)
     Call objCT.TipoBloqueio_KeyPress(KeyAscii)
End Sub

Private Sub TipoBloqueio_Validate(Cancel As Boolean)
     Call objCT.TipoBloqueio_Validate(Cancel)
End Sub

Private Sub DataBloqueio_Change()
     Call objCT.DataBloqueio_Change
End Sub

Private Sub DataBloqueio_GotFocus()
     Call objCT.DataBloqueio_GotFocus
End Sub

Private Sub DataBloqueio_KeyPress(KeyAscii As Integer)
     Call objCT.DataBloqueio_KeyPress(KeyAscii)
End Sub

Private Sub DataBloqueio_Validate(Cancel As Boolean)
     Call objCT.DataBloqueio_Validate(Cancel)
End Sub

Private Sub Observacao_Change()
     Call objCT.Observacao_Change
End Sub

Private Sub Observacao_GotFocus()
     Call objCT.Observacao_GotFocus
End Sub

Private Sub Observacao_KeyPress(KeyAscii As Integer)
     Call objCT.Observacao_KeyPress(KeyAscii)
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)
     Call objCT.Observacao_Validate(Cancel)
End Sub

Private Sub CodUsuario_Change()
     Call objCT.CodUsuario_Change
End Sub

Private Sub CodUsuario_GotFocus()
     Call objCT.CodUsuario_GotFocus
End Sub

Private Sub CodUsuario_KeyPress(KeyAscii As Integer)
     Call objCT.CodUsuario_KeyPress(KeyAscii)
End Sub

Private Sub CodUsuario_Validate(Cancel As Boolean)
     Call objCT.CodUsuario_Validate(Cancel)
End Sub

Private Sub ResponsavelBL_Change()
     Call objCT.ResponsavelBL_Change
End Sub

Private Sub ResponsavelBL_GotFocus()
     Call objCT.ResponsavelBL_GotFocus
End Sub

Private Sub ResponsavelBL_KeyPress(KeyAscii As Integer)
     Call objCT.ResponsavelBL_KeyPress(KeyAscii)
End Sub

Private Sub ResponsavelBL_Validate(Cancel As Boolean)
     Call objCT.ResponsavelBL_Validate(Cancel)
End Sub

Private Sub GridBloqueio_Click()
     Call objCT.GridBloqueio_Click
End Sub

Private Sub GridBloqueio_GotFocus()
     Call objCT.GridBloqueio_GotFocus
End Sub

Private Sub GridBloqueio_EnterCell()
     Call objCT.GridBloqueio_EnterCell
End Sub

Private Sub GridBloqueio_LeaveCell()
     Call objCT.GridBloqueio_LeaveCell
End Sub

Private Sub GridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridBloqueio_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridBloqueio_KeyPress(KeyAscii As Integer)
     Call objCT.GridBloqueio_KeyPress(KeyAscii)
End Sub

Private Sub GridBloqueio_Validate(Cancel As Boolean)
     Call objCT.GridBloqueio_Validate(Cancel)
End Sub

Private Sub GridBloqueio_RowColChange()
     Call objCT.GridBloqueio_RowColChange
End Sub

Private Sub GridBloqueio_Scroll()
     Call objCT.GridBloqueio_Scroll
End Sub

Private Sub GridAlocacao_Click()
     Call objCT.GridAlocacao_Click
End Sub

Private Sub GridAlocacao_GotFocus()
     Call objCT.GridAlocacao_GotFocus
End Sub

Private Sub GridAlocacao_EnterCell()
     Call objCT.GridAlocacao_EnterCell
End Sub

Private Sub GridAlocacao_LeaveCell()
     Call objCT.GridAlocacao_LeaveCell
End Sub

Private Sub GridAlocacao_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridAlocacao_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridAlocacao_KeyPress(KeyAscii As Integer)
     Call objCT.GridAlocacao_KeyPress(KeyAscii)
End Sub

Private Sub GridAlocacao_Validate(Cancel As Boolean)
     Call objCT.GridAlocacao_Validate(Cancel)
End Sub

Private Sub GridAlocacao_RowColChange()
     Call objCT.GridAlocacao_RowColChange
End Sub

Private Sub GridAlocacao_Scroll()
     Call objCT.GridAlocacao_Scroll
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoImprimir_Click()
     Call objCT.BotaoImprimir_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub FilialFaturamento_Validate(Cancel As Boolean)
     Call objCT.FilialFaturamento_Validate(Cancel)
End Sub

Private Sub PercAcrescFin_Validate(Cancel As Boolean)
     Call objCT.PercAcrescFin_Validate(Cancel)
End Sub

Private Sub VolumeEspecie_Change()
     Call objCT.VolumeEspecie_Change
End Sub

Private Sub VolumeMarca_Change()
     Call objCT.VolumeMarca_Change
End Sub

Private Sub VolumeEspecie_Validate(Cancel As Boolean)
     Call objCT.VolumeEspecie_Validate(Cancel)
End Sub

Private Sub VolumeMarca_Validate(Cancel As Boolean)
     Call objCT.VolumeMarca_Validate(Cancel)
End Sub

Private Sub VolumeNumero_Change()
     Call objCT.VolumeNumero_Change
End Sub

Private Sub VolumeQuant_Change()
     Call objCT.VolumeQuant_Change
End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)
     Call objCT.PesoLiquido_Validate(Cancel)
End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)
     Call objCT.PesoBruto_Validate(Cancel)
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

Private Sub ICMSSubstValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstValor1, Source, X, Y)
End Sub

Private Sub ICMSSubstValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstBase1, Source, X, Y)
End Sub

Private Sub ICMSSubstBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstBase1, Button, Shift, X, Y)
End Sub

Private Sub ICMSValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSValor1, Source, X, Y)
End Sub

Private Sub ICMSValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSBase1, Source, X, Y)
End Sub

Private Sub ICMSBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSBase1, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub

Private Sub IPIValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor1, Source, X, Y)
End Sub

Private Sub IPIValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor1, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub MensagemLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MensagemLabel, Source, X, Y)
End Sub

Private Sub MensagemLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MensagemLabel, Button, Shift, X, Y)
End Sub

Private Sub CanalVendaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CanalVendaLabel, Source, X, Y)
End Sub

Private Sub CanalVendaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CanalVendaLabel, Button, Shift, X, Y)
End Sub

Private Sub TranspRedLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TranspRedLabel, Source, X, Y)
End Sub

Private Sub TranspRedLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TranspRedLabel, Button, Shift, X, Y)
End Sub

Private Sub ProdutoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoDescricao, Source, X, Y)
End Sub

Private Sub ProdutoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoDescricao, Button, Shift, X, Y)
End Sub

Private Sub TotalValorComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorComissao, Source, X, Y)
End Sub

Private Sub TotalValorComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorComissao, Button, Shift, X, Y)
End Sub

Private Sub TotalPercentualComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualComissao, Source, X, Y)
End Sub

Private Sub TotalPercentualComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualComissao, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub
'Janaina
Private Sub GridEmb_Click()
     Call objCT.GridEmb_Click
End Sub
'Janaina
Private Sub GridEmb_EnterCell()
     Call objCT.GridEmb_EnterCell
End Sub
'Janaina
Private Sub GridEmb_GotFocus()
     Call objCT.GridEmb_GotFocus
End Sub
'Janaina
Private Sub GridEmb_KeyPress(KeyAscii As Integer)
     Call objCT.GridEmb_KeyPress(KeyAscii)
End Sub
'Janaina
Private Sub GridEmb_LeaveCell()
     Call objCT.GridEmb_LeaveCell
End Sub
'Janaina
Private Sub GridEmb_Validate(Cancel As Boolean)
     Call objCT.GridEmb_Validate(Cancel)
End Sub
'Janaina
Private Sub GridEmb_RowColChange()
     Call objCT.GridEmb_RowColChange
End Sub
'Janaina
Private Sub GridEmb_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridEmb_KeyDown(KeyCode, Shift)
End Sub
'Janaina
Private Sub GridEmb_Scroll()
     Call objCT.GridEmb_Scroll
End Sub
'Janaina
Public Sub ItemEmb_Change()

    Call objCT.ItemEmb_Change

End Sub
'Janaina
Public Sub ItemEmb_GotFocus()

    Call objCT.ItemEmb_GotFocus

End Sub
'Janaina
Public Sub ItemEmb_KeyPress(KeyAscii As Integer)

    Call objCT.ItemEmb_KeyPress(KeyAscii)

End Sub
'Janaina
Public Sub ItemEmb_Validate(Cancel As Boolean)

    Call objCT.ItemEmb_Validate(Cancel)

End Sub
Public Sub Embalagem_Change()

    Call objCT.Embalagem_Change

End Sub

Public Sub Embalagem_GotFocus()

    Call objCT.Embalagem_GotFocus

End Sub

Public Sub Embalagem_KeyPress(KeyAscii As Integer)

    Call objCT.Embalagem_KeyPress(KeyAscii)

End Sub

Public Sub Embalagem_Validate(Cancel As Boolean)

    Call objCT.Embalagem_Validate(Cancel)

End Sub

Public Sub QuantEmb_Change()

    Call objCT.QuantEmb_Change

End Sub

Public Sub QuantEmb_GotFocus()

    Call objCT.QuantEmb_GotFocus

End Sub

Public Sub QuantEmb_KeyPress(KeyAscii As Integer)

    Call objCT.QuantEmb_KeyPress(KeyAscii)

End Sub

Public Sub QuantEmb_Validate(Cancel As Boolean)

    Call objCT.QuantEmb_Validate(Cancel)

End Sub

Public Sub QuantProduto_Change()

    Call objCT.QuantProduto_Change

End Sub

Public Sub QuantProduto_GotFocus()

    Call objCT.QuantProduto_GotFocus

End Sub

Public Sub QuantProduto_KeyPress(KeyAscii As Integer)

    Call objCT.QuantProduto_KeyPress(KeyAscii)

End Sub

Public Sub QuantProduto_Validate(Cancel As Boolean)

    Call objCT.QuantProduto_Validate(Cancel)

End Sub

Private Sub CalculoAuto_Click()

    Call objCT.CalculoAuto_Click

End Sub

Private Sub TotalValorBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorBase, Source, X, Y)
End Sub

Private Sub TotalValorBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorBase, Button, Shift, X, Y)
End Sub

Private Sub TranspRedespacho_Change()
     Call objCT.TranspRedespacho_Change
End Sub

Private Sub TranspRedespacho_Click()
     Call objCT.TranspRedespacho_Click
End Sub

Private Sub TranspRedespacho_Validate(Cancel As Boolean)
     Call objCT.TranspRedespacho_Validate(Cancel)
End Sub

Private Sub Cubagem_Change()
     Call objCT.Cubagem_Change
End Sub

Private Sub Cubagem_Validate(Cancel As Boolean)
    Call objCT.Cubagem_Validate(Cancel)
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub TranspRedLabel_Click()
     Call objCT.TranspRedLabel_Click
End Sub

Private Sub BotaoTodosPedidos_Click()
    Call objCT.BotaoTodosPedidos_Click
End Sub

Private Sub BotaoRefazAlocacao_Click()
    Call objCT.BotaoRefazAlocacao_Click
End Sub

Private Sub BotaoKitVenda_Click()
    Call objCT.BotaoKitVenda_Click
End Sub

Private Sub PedidoRepr_Change()
     Call objCT.PedidoRepr_Change
End Sub

Private Sub PedidoRepr_Validate(Cancel As Boolean)
     Call objCT.PedidoRepr_Validate(Cancel)
End Sub

Private Sub BotaoContato_Click()
     Call objCT.BotaoContato_Click
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

Sub Etapa_Change()
     Call objCT.Projeto_Change
End Sub

Sub Etapa_Click()
     Call objCT.Projeto_Change
End Sub

Sub Etapa_Validate(Cancel As Boolean)
     Call objCT.Projeto_Validate(Cancel)
End Sub

Private Sub TipoPagto_Change()
     Call objCT.TipoPagto_Change
End Sub

Private Sub TipoPagto_GotFocus()
     Call objCT.TipoPagto_GotFocus
End Sub

Private Sub TipoPagto_KeyPress(KeyAscii As Integer)
     Call objCT.TipoPagto_KeyPress(KeyAscii)
End Sub

Private Sub TipoPagto_Validate(Cancel As Boolean)
     Call objCT.TipoPagto_Validate(Cancel)
End Sub

Private Sub BotaoTipoPagto_Click()
     Call objCT.BotaoTipoPagto_Click
End Sub

Private Sub BotaoEntrega_Click()
    Call objCT.BotaoEntrega_Click
End Sub

Private Sub BotaoDataRefFluxoDown_Click()
     Call objCT.BotaoDataRefFluxoDown_Click
End Sub

Private Sub BotaoDataRefFluxoUp_Click()
     Call objCT.BotaoDataRefFluxoUp_Click
End Sub

Private Sub DataRefFluxo_GotFocus()
     Call objCT.DataRefFluxo_GotFocus
End Sub

Private Sub DataRefFluxo_Change()
     Call objCT.DataRefFluxo_Change
End Sub

Private Sub DataRefFluxo_Validate(Cancel As Boolean)
     Call objCT.DataRefFluxo_Validate(Cancel)
End Sub

Private Sub RecebForaFaixa_Change()
     Call objCT.RecebForaFaixa_Change
End Sub

Private Sub RecebForaFaixa_Click()
     Call objCT.RecebForaFaixa_Click
End Sub

Private Sub RecebForaFaixa_GotFocus()
     Call objCT.RecebForaFaixa_GotFocus
End Sub

Private Sub RecebForaFaixa_KeyPress(KeyAscii As Integer)
     Call objCT.RecebForaFaixa_KeyPress(KeyAscii)
End Sub

Private Sub RecebForaFaixa_Validate(Cancel As Boolean)
     Call objCT.RecebForaFaixa_Validate(Cancel)
End Sub

Private Sub PercentMaisReceb_Change()
     Call objCT.PercentMaisReceb_Change
End Sub

Private Sub PercentMaisReceb_GotFocus()
     Call objCT.PercentMaisReceb_GotFocus
End Sub

Private Sub PercentMaisReceb_KeyPress(KeyAscii As Integer)
     Call objCT.PercentMaisReceb_KeyPress(KeyAscii)
End Sub

Private Sub PercentMaisReceb_Validate(Cancel As Boolean)
     Call objCT.PercentMaisReceb_Validate(Cancel)
End Sub

Private Sub PercentMenosReceb_Change()
     Call objCT.PercentMenosReceb_Change
End Sub

Private Sub PercentMenosReceb_GotFocus()
     Call objCT.PercentMenosReceb_GotFocus
End Sub

Private Sub PercentMenosReceb_KeyPress(KeyAscii As Integer)
     Call objCT.PercentMenosReceb_KeyPress(KeyAscii)
End Sub

Private Sub PercentMenosReceb_Validate(Cancel As Boolean)
     Call objCT.PercentMenosReceb_Validate(Cancel)
End Sub

Private Sub QuantFatAMais_Change()
     Call objCT.QuantFatAMais_Change
End Sub

Private Sub QuantFatAMais_GotFocus()
     Call objCT.QuantFatAMais_GotFocus
End Sub

Private Sub QuantFatAMais_KeyPress(KeyAscii As Integer)
     Call objCT.QuantFatAMais_KeyPress(KeyAscii)
End Sub

Private Sub QuantFatAMais_Validate(Cancel As Boolean)
     Call objCT.QuantFatAMais_Validate(Cancel)
End Sub

Private Sub Prioridade_GotFocus()
     Call objCT.Prioridade_GotFocus
End Sub

Private Sub Prioridade_KeyPress(KeyAscii As Integer)
     Call objCT.Prioridade_KeyPress(KeyAscii)
End Sub

Private Sub Prioridade_Validate(Cancel As Boolean)
     Call objCT.Prioridade_Validate(Cancel)
End Sub

Private Sub PrioridadePadrao_GotFocus()
     Call objCT.PrioridadePadrao_GotFocus
End Sub

Public Sub BotaoEmail_Click()
     Call objCT.BotaoEmail_Click
End Sub

Private Sub ValorDescontoItens_Change()
     Call objCT.ValorDescontoItens_Change
End Sub

Private Sub ValorDescontoItens_Validate(Cancel As Boolean)
     Call objCT.ValorDescontoItens_Validate(Cancel)
End Sub

Private Sub PercDescontoItens_Change()
     Call objCT.PercDescontoItens_Change
End Sub

Private Sub PercDescontoItens_Validate(Cancel As Boolean)
     Call objCT.PercDescontoItens_Validate(Cancel)
End Sub

Private Sub ValorDescontoTit_Change()
     Call objCT.ValorDescontoTit_Change
End Sub

Private Sub ValorDescontoTit_Validate(Cancel As Boolean)
     Call objCT.ValorDescontoTit_Validate(Cancel)
End Sub

Private Sub BotaoInfoAdic_Click()
     Call objCT.BotaoInfoAdic_Click
End Sub

Private Sub BotaoImprimirConf_Click()
     Call objCT.BotaoImprimirConf_Click
End Sub

Private Sub Parc_Change()
     Call objCT.Parc_Change
End Sub

Private Sub Parc_GotFocus()
     Call objCT.Parc_GotFocus
End Sub

Private Sub TabPrecoItemPV_Change()
     Call objCT.TabPrecoItemPV_Change
End Sub

Private Sub TabPrecoItemPV_Click()
     Call objCT.TabPrecoItemPV_Click
End Sub

Private Sub TabPrecoItemPV_GotFocus()
     Call objCT.TabPrecoItemPV_GotFocus
End Sub

Private Sub TabPrecoItemPV_KeyPress(KeyAscii As Integer)
     Call objCT.TabPrecoItemPV_KeyPress(KeyAscii)
End Sub

Private Sub TabPrecoItemPV_Validate(Cancel As Boolean)
     Call objCT.TabPrecoItemPV_Validate(Cancel)
End Sub

Private Sub ComissaoItemPV_Change()
     Call objCT.ComissaoItemPV_Change
End Sub

Private Sub ComissaoItemPV_GotFocus()
     Call objCT.ComissaoItemPV_GotFocus
End Sub

Private Sub ComissaoItemPV_KeyPress(KeyAscii As Integer)
     Call objCT.ComissaoItemPV_KeyPress(KeyAscii)
End Sub

Private Sub ComissaoItemPV_Validate(Cancel As Boolean)
     Call objCT.ComissaoItemPV_Validate(Cancel)
End Sub



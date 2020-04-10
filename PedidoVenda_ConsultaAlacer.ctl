VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PedidoVenda_Consulta 
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9540
   ScaleHeight     =   6375
   ScaleWidth      =   9540
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   3
      Left            =   105
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   9270
      Begin VB.Frame Frame8 
         Caption         =   "Volumes"
         Height          =   735
         Left            =   240
         TabIndex        =   113
         Top             =   1980
         Width           =   8835
         Begin VB.Label VolumeMarca 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5355
            TabIndex        =   146
            Top             =   285
            Width           =   1020
         End
         Begin VB.Label VolumeNumero 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7140
            TabIndex        =   147
            Top             =   285
            Width           =   1440
         End
         Begin VB.Label VolumeEspecie 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3090
            TabIndex        =   148
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label VolumeQuant 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1395
            TabIndex        =   149
            Top             =   278
            Width           =   585
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
            Left            =   4695
            TabIndex        =   150
            Top             =   330
            Width           =   600
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
            Left            =   2295
            TabIndex        =   151
            Top             =   330
            Width           =   750
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
            Left            =   300
            TabIndex        =   152
            Top             =   330
            Width           =   1050
         End
         Begin VB.Label Label44 
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
            Index           =   8
            Left            =   6750
            TabIndex        =   153
            Top             =   330
            Width           =   345
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Complemento"
         Height          =   1740
         Index           =   8
         Left            =   225
         TabIndex        =   114
         Top             =   2775
         Width           =   8820
         Begin VB.TextBox Mensagem 
            BackColor       =   &H8000000F&
            Height          =   750
            Left            =   1485
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   212
            Top             =   180
            Width           =   7215
         End
         Begin MSMask.MaskEdBox Cubagem 
            Height          =   300
            Left            =   6825
            TabIndex        =   194
            Top             =   1020
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483644
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ped. Repres.:"
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
            Left            =   5610
            TabIndex        =   211
            Top             =   1440
            Width           =   1185
         End
         Begin VB.Label PedidoRepr 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6825
            TabIndex        =   210
            Top             =   1395
            Width           =   1890
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
            Left            =   60
            TabIndex        =   197
            Top             =   1455
            Width           =   1425
         End
         Begin VB.Label CanalVenda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1485
            TabIndex        =   196
            Top             =   1395
            Width           =   1365
         End
         Begin VB.Label Label30 
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
            Index           =   7
            Left            =   5940
            TabIndex        =   195
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label PedidoCliente 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4095
            TabIndex        =   154
            Top             =   1395
            Width           =   1365
         End
         Begin VB.Label PesoLiquido 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1485
            TabIndex        =   155
            Top             =   1035
            Width           =   1365
         End
         Begin VB.Label PesoBruto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4095
            TabIndex        =   156
            Top             =   1035
            Width           =   1365
         End
         Begin VB.Label Label30 
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
            Index           =   6
            Left            =   3075
            TabIndex        =   157
            Top             =   1065
            Width           =   1005
         End
         Begin VB.Label Label30 
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
            Index           =   5
            Left            =   270
            TabIndex        =   158
            Top             =   1080
            Width           =   1215
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
            TabIndex        =   159
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ped. Cliente:"
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
            Left            =   2970
            TabIndex        =   160
            Top             =   1425
            Width           =   1110
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Dados de Entrega"
         Height          =   1920
         Index           =   6
         Left            =   225
         TabIndex        =   115
         Top             =   0
         Width           =   8820
         Begin VB.Frame Frame16 
            Caption         =   "Redespacho"
            Height          =   990
            Left            =   4860
            TabIndex        =   187
            Top             =   570
            Width           =   3870
            Begin VB.CheckBox RedespachoCli 
               Caption         =   "por conta do cliente"
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
               Height          =   270
               Left            =   225
               TabIndex        =   188
               Top             =   615
               Width           =   2100
            End
            Begin VB.Label TranspRedespacho 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1515
               TabIndex        =   190
               Top             =   180
               Width           =   2235
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
               TabIndex        =   189
               Top             =   255
               Width           =   1365
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frete por conta"
            Height          =   795
            Left            =   240
            TabIndex        =   116
            Top             =   690
            Width           =   1575
            Begin VB.Label TipoFrete 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   120
               TabIndex        =   206
               Top             =   345
               Width           =   1305
            End
         End
         Begin VB.Label DataEntregaReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6375
            TabIndex        =   215
            Top             =   1560
            Width           =   1185
         End
         Begin VB.Label LabelVLight 
            AutoSize        =   -1  'True
            Caption         =   "Entrega Real:"
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
            Left            =   5115
            TabIndex        =   214
            Top             =   1620
            Width           =   1185
         End
         Begin VB.Label LabelVLight 
            AutoSize        =   -1  'True
            Caption         =   "Entrega Prevista:"
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
            Left            =   450
            TabIndex        =   192
            Top             =   1590
            Width           =   1485
         End
         Begin VB.Label DataEntregaPV 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1950
            TabIndex        =   191
            Top             =   1530
            Width           =   1185
         End
         Begin VB.Label Placa 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3315
            TabIndex        =   161
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label PlacaUF 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3315
            TabIndex        =   162
            Top             =   1140
            Width           =   735
         End
         Begin VB.Label Transportadora 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6360
            TabIndex        =   163
            Top             =   210
            Width           =   2235
         End
         Begin VB.Label FilialEntrega 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1995
            TabIndex        =   164
            Top             =   285
            Width           =   1935
         End
         Begin VB.Label Label30 
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
            Index           =   3
            Left            =   2025
            TabIndex        =   165
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label30 
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
            Index           =   4
            Left            =   1995
            TabIndex        =   166
            Top             =   780
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
            Left            =   4920
            TabIndex        =   167
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label LabelVLight 
            AutoSize        =   -1  'True
            Caption         =   "Filial Entrega:"
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
            Left            =   735
            TabIndex        =   168
            Top             =   345
            Width           =   1185
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4530
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   9270
      Begin VB.CommandButton Command1 
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
         Left            =   180
         TabIndex        =   186
         Top             =   3765
         Width           =   1365
      End
      Begin VB.Frame Frame12 
         Caption         =   "Valores"
         Height          =   1230
         Index           =   4
         Left            =   180
         TabIndex        =   117
         Top             =   2490
         Width           =   8865
         Begin VB.Label ISSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7965
            TabIndex        =   207
            Top             =   855
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   450
            TabIndex        =   198
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label ValorDesconto 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   450
            TabIndex        =   123
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label ValorDespesas 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4380
            TabIndex        =   124
            Top             =   885
            Width           =   1500
         End
         Begin VB.Label ValorSeguro 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3045
            TabIndex        =   125
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label ValorFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1755
            TabIndex        =   126
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7380
            TabIndex        =   127
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6060
            TabIndex        =   128
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7365
            TabIndex        =   129
            Top             =   375
            Width           =   1125
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1770
            TabIndex        =   130
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3030
            TabIndex        =   131
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4395
            TabIndex        =   132
            Top             =   390
            Width           =   1500
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6060
            TabIndex        =   133
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Desconto        Base ICMS          ICMS         Base ICMS Subst    ICMS Subst       Produtos"
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
            Left            =   570
            TabIndex        =   134
            Top             =   180
            Width           =   7695
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "ISS                 Frete             Seguro              Despesas               IPI                Total"
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
            Left            =   720
            TabIndex        =   135
            Top             =   675
            Width           =   7350
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Itens"
         Height          =   2415
         Index           =   3
         Left            =   180
         TabIndex        =   118
         Top             =   60
         Width           =   8865
         Begin MSMask.MaskEdBox ComissaoItemPV 
            Height          =   255
            Left            =   6795
            TabIndex        =   209
            Top             =   1530
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
            ItemData        =   "PedidoVenda_ConsultaAlacer.ctx":0000
            Left            =   5085
            List            =   "PedidoVenda_ConsultaAlacer.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   1470
            Width           =   1800
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "PedidoVenda_ConsultaAlacer.ctx":0004
            Left            =   1500
            List            =   "PedidoVenda_ConsultaAlacer.ctx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   390
            Width           =   720
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3780
            MaxLength       =   50
            TabIndex        =   14
            Top             =   900
            Width           =   2085
         End
         Begin MSMask.MaskEdBox QuantCancelada 
            Height          =   225
            Left            =   7065
            TabIndex        =   16
            Top             =   510
            Width           =   1485
            _ExtentX        =   2619
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
         Begin MSMask.MaskEdBox QuantFaturada 
            Height          =   225
            Left            =   6900
            TabIndex        =   18
            Top             =   870
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
            Left            =   5325
            TabIndex        =   17
            Top             =   870
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
            Left            =   2565
            TabIndex        =   13
            Top             =   810
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
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   1335
            TabIndex        =   12
            Top             =   780
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
            Left            =   180
            TabIndex        =   11
            Top             =   810
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   4110
            TabIndex        =   10
            Top             =   540
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2505
            TabIndex        =   9
            Top             =   450
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   270
            TabIndex        =   7
            Top             =   435
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   5595
            TabIndex        =   15
            Top             =   510
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
            Height          =   1455
            Left            =   135
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   330
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   2566
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
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos..."
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
         Left            =   7680
         TabIndex        =   20
         Top             =   3780
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4605
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   1560
      Width           =   9255
      Begin VB.Frame Frame1 
         Caption         =   "Projetos"
         Height          =   1080
         Index           =   0
         Left            =   270
         TabIndex        =   200
         Top             =   3030
         Width           =   8865
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
            Left            =   3105
            TabIndex        =   202
            Top             =   405
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   201
            Top             =   405
            Width           =   2550
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   1230
            TabIndex        =   203
            Top             =   420
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483633
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   495
            TabIndex        =   205
            Top             =   465
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
            Left            =   4170
            TabIndex        =   204
            Top             =   465
            Width           =   570
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Identificação"
         Height          =   1350
         Index           =   0
         Left            =   240
         TabIndex        =   87
         Top             =   210
         Width           =   8925
         Begin VB.CommandButton BotaoDesfazer 
            Caption         =   "Desfazer Baixa"
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
            Height          =   735
            Left            =   7590
            TabIndex        =   199
            Top             =   345
            Width           =   1215
         End
         Begin VB.Label LabelVLight 
            AutoSize        =   -1  'True
            Caption         =   "Filial do Pedido:"
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
            Left            =   3750
            TabIndex        =   177
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label FilialPedido 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5160
            TabIndex        =   176
            Top             =   360
            Width           =   2145
         End
         Begin VB.Label FilialFaturamento 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5160
            TabIndex        =   136
            Top             =   810
            Width           =   2145
         End
         Begin VB.Label NatOpInternaEspelho 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1650
            TabIndex        =   137
            Top             =   570
            Width           =   480
         End
         Begin VB.Label LabelVLight 
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
            Height          =   195
            Index           =   2
            Left            =   3555
            TabIndex        =   138
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label NaturezaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nat. Operação:"
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
            TabIndex        =   139
            Top             =   600
            Width           =   1320
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Preços"
         Height          =   1050
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   1710
         Width           =   8925
         Begin VB.Label PercAcrescFin 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4800
            TabIndex        =   140
            Top             =   420
            Width           =   765
         End
         Begin VB.Label CondicaoPagamento 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1275
            TabIndex        =   141
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label TabelaPreco 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6930
            TabIndex        =   142
            Top             =   420
            Width           =   1875
         End
         Begin VB.Label LabelVLight 
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
            Index           =   1
            Left            =   5700
            TabIndex        =   143
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label LabelVLight 
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
            Index           =   5
            Left            =   3270
            TabIndex        =   144
            Top             =   480
            Width           =   1485
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
            Left            =   180
            TabIndex        =   145
            Top             =   480
            Width           =   1065
         End
      End
      Begin VB.CheckBox FaturaIntegral 
         Caption         =   "Só libera o pedido integralmente"
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
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   4305
         Width           =   3300
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4650
      Index           =   9
      Left            =   135
      TabIndex        =   81
      Top             =   1560
      Visible         =   0   'False
      Width           =   9270
      Begin TelasFATAla.TabTributacaoFat TabTrib 
         Height          =   4560
         Left            =   90
         TabIndex        =   213
         Top             =   0
         Width           =   9000
         _extentx        =   15875
         _extenty        =   8043
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4575
      Index           =   4
      Left            =   120
      TabIndex        =   89
      Top             =   1560
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   3795
         Left            =   60
         TabIndex        =   112
         Top             =   225
         Width           =   9165
         Begin VB.ComboBox TipoDesconto3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2985
            TabIndex        =   27
            Top             =   1950
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   3030
            TabIndex        =   26
            Top             =   1620
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2985
            TabIndex        =   25
            Top             =   1320
            Width           =   1965
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7380
            TabIndex        =   34
            Top             =   1365
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox Desconto3Valor 
            Height          =   225
            Left            =   6015
            TabIndex        =   33
            Top             =   2010
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Ate 
            Height          =   225
            Left            =   4905
            TabIndex        =   30
            Top             =   2010
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
         Begin MSMask.MaskEdBox Desconto2Valor 
            Height          =   225
            Left            =   6045
            TabIndex        =   32
            Top             =   1695
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Ate 
            Height          =   225
            Left            =   4905
            TabIndex        =   29
            Top             =   1695
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
         Begin MSMask.MaskEdBox Desconto1Valor 
            Height          =   225
            Left            =   6030
            TabIndex        =   31
            Top             =   1365
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto1Ate 
            Height          =   225
            Left            =   4890
            TabIndex        =   28
            Top             =   1365
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
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   480
            TabIndex        =   23
            Top             =   1335
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
            Height          =   240
            Left            =   1605
            TabIndex        =   24
            Top             =   1350
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Percentual 
            Height          =   225
            Left            =   7410
            TabIndex        =   35
            Top             =   1710
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox Desconto3Percentual 
            Height          =   225
            Left            =   7365
            TabIndex        =   36
            Top             =   2010
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2895
            Left            =   60
            TabIndex        =   37
            Top             =   360
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   50
            Cols            =   6
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
      Height          =   4125
      Index           =   6
      Left            =   120
      TabIndex        =   50
      Top             =   1560
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame SSFrame1 
         Caption         =   "Bloqueios"
         Height          =   3630
         Left            =   75
         TabIndex        =   109
         Top             =   135
         Width           =   9120
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   1020
            MaxLength       =   250
            TabIndex        =   193
            Top             =   1320
            Width           =   4245
         End
         Begin VB.ComboBox TipoBloqueio 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "PedidoVenda_ConsultaAlacer.ctx":0008
            Left            =   180
            List            =   "PedidoVenda_ConsultaAlacer.ctx":000A
            TabIndex        =   51
            Top             =   360
            Width           =   1605
         End
         Begin MSMask.MaskEdBox ResponsavelLib 
            Height          =   270
            Left            =   7005
            TabIndex        =   56
            Top             =   375
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
            Left            =   5775
            TabIndex        =   55
            Top             =   375
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
            Left            =   2955
            TabIndex        =   53
            Top             =   375
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
            Left            =   4410
            TabIndex        =   54
            Top             =   375
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
         Begin MSMask.MaskEdBox DataBloqueio 
            Height          =   270
            Left            =   1815
            TabIndex        =   52
            Top             =   375
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
            Left            =   150
            TabIndex        =   110
            Top             =   270
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4155
      Index           =   5
      Left            =   120
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   9240
      Begin VB.CommandButton BotaoVendedores 
         Caption         =   "Vendedores ..."
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
         Left            =   7170
         TabIndex        =   49
         Top             =   3450
         Width           =   1785
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
         Left            =   1050
         TabIndex        =   39
         Top             =   60
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3360
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   3795
         Index           =   0
         Left            =   60
         TabIndex        =   111
         Top             =   270
         Width           =   9060
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
            Height          =   645
            Index           =   1
            Left            =   120
            TabIndex        =   179
            Top             =   3090
            Width           =   6885
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
               TabIndex        =   185
               Top             =   360
               Width           =   1095
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
               TabIndex        =   184
               Top             =   360
               Width           =   615
            End
            Begin VB.Label TotalPercentualComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3960
               TabIndex        =   183
               Top             =   345
               Width           =   825
            End
            Begin VB.Label TotalValorComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5640
               TabIndex        =   182
               Top             =   345
               Width           =   1155
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
               Index           =   5
               Left            =   360
               TabIndex        =   181
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label TotalValorBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1440
               TabIndex        =   180
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.ComboBox DiretoIndireto 
            Height          =   315
            ItemData        =   "PedidoVenda_ConsultaAlacer.ctx":000C
            Left            =   6120
            List            =   "PedidoVenda_ConsultaAlacer.ctx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   178
            Top             =   960
            Width           =   1335
         End
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   225
            Left            =   3525
            TabIndex        =   43
            Top             =   330
            Width           =   1065
            _ExtentX        =   1879
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   225
            Left            =   2400
            TabIndex        =   42
            Top             =   345
            Width           =   1065
            _ExtentX        =   1879
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   225
            Left            =   1515
            TabIndex        =   41
            Top             =   525
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   397
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
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   225
            Left            =   120
            TabIndex        =   40
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   225
            Left            =   5580
            TabIndex        =   45
            Top             =   345
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
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   225
            Left            =   4725
            TabIndex        =   44
            Top             =   345
            Width           =   840
            _ExtentX        =   1482
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
         Begin MSMask.MaskEdBox PercentualBaixa 
            Height          =   225
            Left            =   6720
            TabIndex        =   46
            Top             =   345
            Width           =   825
            _ExtentX        =   1455
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
         Begin MSMask.MaskEdBox ValorBaixa 
            Height          =   225
            Left            =   7575
            TabIndex        =   47
            Top             =   345
            Width           =   1065
            _ExtentX        =   1879
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
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   2790
            Left            =   135
            TabIndex        =   48
            Top             =   240
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   4921
            _Version        =   393216
            Rows            =   11
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
      Caption         =   "Frame14"
      Height          =   4125
      Index           =   8
      Left            =   135
      TabIndex        =   67
      Top             =   1560
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame FrameNFiscal 
         Caption         =   "Ordenado por Série + Nota Fiscal"
         Height          =   2775
         Index           =   2
         Left            =   150
         TabIndex        =   88
         Top             =   570
         Visible         =   0   'False
         Width           =   8955
         Begin VB.TextBox SerieNF 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3600
            TabIndex        =   76
            Top             =   1200
            Width           =   795
         End
         Begin VB.TextBox ItemNF 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   930
            TabIndex        =   69
            Top             =   420
            Width           =   795
         End
         Begin VB.TextBox NFiscalNF 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   300
            TabIndex        =   70
            Top             =   420
            Width           =   765
         End
         Begin VB.TextBox DescricaoProdutoNF 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3510
            MaxLength       =   50
            TabIndex        =   72
            Top             =   450
            Width           =   2115
         End
         Begin VB.TextBox UMNF 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4800
            TabIndex        =   90
            Top             =   450
            Width           =   645
         End
         Begin VB.TextBox ItemPVNF 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1560
            TabIndex        =   91
            Top             =   420
            Width           =   855
         End
         Begin MSMask.MaskEdBox PrecoTotalNF 
            Height          =   225
            Left            =   2340
            TabIndex        =   75
            Top             =   1200
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
         Begin MSMask.MaskEdBox DescontoNF 
            Height          =   225
            Left            =   1200
            TabIndex        =   74
            Top             =   1170
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
         Begin MSMask.MaskEdBox PercDescNF 
            Height          =   225
            Left            =   210
            TabIndex        =   73
            Top             =   1170
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox PrecoUnitarioNF 
            Height          =   225
            Left            =   6960
            TabIndex        =   77
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox QuantidadeNF 
            Height          =   225
            Left            =   5430
            TabIndex        =   78
            Top             =   450
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
         Begin MSMask.MaskEdBox ProdutoNF 
            Height          =   225
            Left            =   2370
            TabIndex        =   71
            Top             =   420
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridNFiscal 
            Height          =   2265
            Left            =   210
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   330
            Width           =   8625
            _ExtentX        =   15214
            _ExtentY        =   3995
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
      Begin VB.Frame FrameNFiscal 
         Caption         =   "Ordenado por Item Pedido de Venda"
         Height          =   2775
         Index           =   1
         Left            =   150
         TabIndex        =   92
         Top             =   570
         Width           =   8955
         Begin VB.TextBox SeriePV 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4350
            TabIndex        =   93
            Top             =   750
            Width           =   795
         End
         Begin MSMask.MaskEdBox PrecoTotalPV 
            Height          =   225
            Left            =   3540
            TabIndex        =   94
            Top             =   2310
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
         Begin MSMask.MaskEdBox DescontoPV 
            Height          =   225
            Left            =   2550
            TabIndex        =   95
            Top             =   2310
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
         Begin MSMask.MaskEdBox PercDescPV 
            Height          =   225
            Left            =   1590
            TabIndex        =   96
            Top             =   2340
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox PrecoUnitarioPV 
            Height          =   225
            Left            =   210
            TabIndex        =   97
            Top             =   2340
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox QuantidadePV 
            Height          =   225
            Left            =   6870
            TabIndex        =   98
            Top             =   480
            Width           =   1320
            _ExtentX        =   2328
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
         Begin VB.TextBox ItemNFPV 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6240
            TabIndex        =   99
            Top             =   480
            Width           =   765
         End
         Begin VB.TextBox NFiscalPV 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   5610
            TabIndex        =   100
            Top             =   810
            Width           =   795
         End
         Begin MSMask.MaskEdBox QuantFaturadaPV 
            Height          =   225
            Left            =   4110
            TabIndex        =   101
            Top             =   510
            Width           =   1410
            _ExtentX        =   2487
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
         Begin VB.TextBox DescricaoProdutoPV 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2190
            MaxLength       =   50
            TabIndex        =   102
            Top             =   750
            Width           =   2115
         End
         Begin VB.TextBox UMPV 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3480
            TabIndex        =   103
            Top             =   450
            Width           =   675
         End
         Begin VB.TextBox ItemPV 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   240
            TabIndex        =   104
            Top             =   450
            Width           =   885
         End
         Begin MSMask.MaskEdBox ProdutoPV 
            Height          =   225
            Left            =   1050
            TabIndex        =   105
            Top             =   465
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridNFItemPV 
            Height          =   2265
            Left            =   210
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   330
            Width           =   8625
            _ExtentX        =   15214
            _ExtentY        =   3995
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
      Begin VB.CommandButton BotaoNFiscal 
         Caption         =   "Nota Fiscal ..."
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
         Left            =   7230
         TabIndex        =   80
         Top             =   3540
         Width           =   1875
      End
      Begin VB.ComboBox ComboOrdenacao 
         Height          =   315
         ItemData        =   "PedidoVenda_ConsultaAlacer.ctx":002C
         Left            =   2580
         List            =   "PedidoVenda_ConsultaAlacer.ctx":0036
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   105
         Width           =   2475
      End
      Begin VB.Label Label44 
         Caption         =   "Ordenados por:"
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
         Index           =   9
         Left            =   1080
         TabIndex        =   169
         Top             =   150
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4125
      Index           =   7
      Left            =   120
      TabIndex        =   57
      Top             =   1560
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame12 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   735
         Index           =   5
         Left            =   105
         TabIndex        =   107
         Top             =   3330
         Width           =   5790
         Begin VB.Label ProdutoDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1140
            TabIndex        =   121
            Top             =   300
            Width           =   4395
         End
         Begin VB.Label Label8 
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
            Index           =   4
            Left            =   240
            TabIndex        =   122
            Top             =   330
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Reserva dos Produtos"
         Height          =   3300
         Left            =   105
         TabIndex        =   108
         Top             =   -15
         Width           =   8940
         Begin VB.TextBox Responsavel 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5625
            TabIndex        =   63
            Top             =   1035
            Width           =   2115
         End
         Begin MSMask.MaskEdBox UnidadeMedEst 
            Height          =   225
            Left            =   7830
            TabIndex        =   65
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
            Left            =   6510
            TabIndex        =   64
            Top             =   585
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
            Left            =   1170
            TabIndex        =   59
            Top             =   630
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
            Left            =   2385
            TabIndex        =   60
            Top             =   615
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
            Left            =   3660
            TabIndex        =   61
            Top             =   630
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
            Left            =   5070
            TabIndex        =   62
            Top             =   600
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
            Left            =   600
            TabIndex        =   58
            Top             =   630
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
         Begin MSFlexGridLib.MSFlexGrid GridReserva 
            Height          =   2805
            Left            =   180
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   225
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
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   6510
      ScaleHeight     =   795
      ScaleWidth      =   2805
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   30
      Width           =   2865
      Begin VB.CommandButton BotaoConsulta 
         Caption         =   "Consultar"
         Height          =   735
         Left            =   60
         Picture         =   "PedidoVenda_ConsultaAlacer.ctx":005A
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Consultar"
         Top             =   30
         Width           =   825
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   735
         Left            =   1725
         Picture         =   "PedidoVenda_ConsultaAlacer.ctx":0B28
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Limpar"
         Top             =   30
         Width           =   480
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   735
         Left            =   2265
         Picture         =   "PedidoVenda_ConsultaAlacer.ctx":105A
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Fechar"
         Top             =   30
         Width           =   480
      End
      Begin VB.CommandButton BotaoEditar 
         Caption         =   "Editar"
         Height          =   735
         Left            =   945
         Picture         =   "PedidoVenda_ConsultaAlacer.ctx":11D8
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Editar"
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4140
      TabIndex        =   1
      Top             =   120
      Width           =   2145
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   975
      TabIndex        =   0
      Top             =   135
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   3540
      TabIndex        =   3
      Top             =   540
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownEmissao 
      Height          =   300
      Left            =   2085
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   540
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataEmissao 
      Height          =   300
      Left            =   975
      TabIndex        =   2
      Top             =   540
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5295
      Left            =   90
      TabIndex        =   119
      Top             =   960
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   9340
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
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
            Caption         =   "Notas Fiscais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
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
   Begin VB.Label StatusPedido 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5505
      TabIndex        =   170
      Top             =   540
      Width           =   780
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
      Index           =   1
      Left            =   4860
      TabIndex        =   171
      Top             =   593
      Width           =   615
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
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   172
      Top             =   585
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
      Height          =   195
      Left            =   2790
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   173
      Top             =   600
      Width           =   720
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
      Left            =   300
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   174
      Top             =   180
      Width           =   660
   End
   Begin VB.Label LabelVLight 
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
      Index           =   0
      Left            =   3645
      TabIndex        =   175
      Top             =   180
      Width           =   465
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9540
      Y1              =   930
      Y2              =   930
   End
End
Attribute VB_Name = "PedidoVenda_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Event Unload()

Private WithEvents objCT As CTPV_Consulta
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTPV_Consulta
    Set objCT.objUserControl = Me
    Set objCT.gobjInfoUsu = New CTPedVendCVGAla
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTPedVendCAla
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoConsulta_Click()
     Call objCT.BotaoConsulta_Click
End Sub

Private Sub BotaoEditar_Click()
     Call objCT.BotaoEditar_Click
End Sub

Private Sub BotaoNFiscal_Click()
     Call objCT.BotaoNFiscal_Click
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub ComboOrdenacao_Click()
     Call objCT.ComboOrdenacao_Click
End Sub

Private Sub Command1_Click()
     Call objCT.Command1_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub LabelCliente_Click()
     Call objCT.LabelCliente_Click
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Function Trata_Parametros(Optional objPedidoVenda As ClassPedidoDeVenda) As Long
     Trata_Parametros = objCT.Trata_Parametros(objPedidoVenda)
End Function

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
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

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub BotaoVendedores_Click()
     Call objCT.BotaoVendedores_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub GridReserva_RowColChange()
     Call objCT.GridReserva_RowColChange
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
     Call objCT.Opcao_BeforeClick(Cancel)
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

Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label13(Index), Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13(Index), Button, Shift, X, Y)
End Sub

Private Sub Label44_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label44(Index), Source, X, Y)
End Sub

Private Sub Label44_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label44(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelVLight_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelVLight(Index), Source, X, Y)
End Sub

Private Sub LabelVLight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVLight(Index), Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label30(Index), Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30(Index), Button, Shift, X, Y)
End Sub

Private Sub ProdutoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoDescricao, Source, X, Y)
End Sub

Private Sub ProdutoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoDescricao, Button, Shift, X, Y)
End Sub

Private Sub ValorDesconto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDesconto, Source, X, Y)
End Sub

Private Sub ValorDesconto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDesconto, Button, Shift, X, Y)
End Sub

Private Sub ValorDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDespesas, Source, X, Y)
End Sub

Private Sub ValorDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDespesas, Button, Shift, X, Y)
End Sub

Private Sub ValorSeguro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorSeguro, Source, X, Y)
End Sub

Private Sub ValorSeguro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorSeguro, Button, Shift, X, Y)
End Sub

Private Sub ValorFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorFrete, Source, X, Y)
End Sub

Private Sub ValorFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorFrete, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub IPIValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor1, Source, X, Y)
End Sub

Private Sub IPIValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor1, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
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

Private Sub FilialFaturamento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialFaturamento, Source, X, Y)
End Sub

Private Sub FilialFaturamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialFaturamento, Button, Shift, X, Y)
End Sub

'Private Sub NaturezaOp_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(NaturezaOp, Source, X, Y)
'End Sub
'
'Private Sub NaturezaOp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(NaturezaOp, Button, Shift, X, Y)
'End Sub
'
'Private Sub NaturezaLabel_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(NaturezaLabel, Source, X, Y)
'End Sub
'
'Private Sub NaturezaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(NaturezaLabel, Button, Shift, X, Y)
'End Sub

Private Sub PercAcrescFin_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercAcrescFin, Source, X, Y)
End Sub

Private Sub PercAcrescFin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercAcrescFin, Button, Shift, X, Y)
End Sub

Private Sub CondicaoPagamento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagamento, Source, X, Y)
End Sub

Private Sub CondicaoPagamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagamento, Button, Shift, X, Y)
End Sub

Private Sub TabelaPreco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TabelaPreco, Source, X, Y)
End Sub

Private Sub TabelaPreco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TabelaPreco, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub VolumeMarca_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(VolumeMarca, Source, X, Y)
End Sub

Private Sub VolumeMarca_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VolumeMarca, Button, Shift, X, Y)
End Sub

Private Sub VolumeNumero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(VolumeNumero, Source, X, Y)
End Sub

Private Sub VolumeNumero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VolumeNumero, Button, Shift, X, Y)
End Sub

Private Sub VolumeEspecie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(VolumeEspecie, Source, X, Y)
End Sub

Private Sub VolumeEspecie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VolumeEspecie, Button, Shift, X, Y)
End Sub

Private Sub VolumeQuant_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(VolumeQuant, Source, X, Y)
End Sub

Private Sub VolumeQuant_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VolumeQuant, Button, Shift, X, Y)
End Sub

Private Sub PedidoCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PedidoCliente, Source, X, Y)
End Sub

Private Sub PedidoCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PedidoCliente, Button, Shift, X, Y)
End Sub

Private Sub PesoLiquido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PesoLiquido, Source, X, Y)
End Sub

Private Sub PesoLiquido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PesoLiquido, Button, Shift, X, Y)
End Sub

Private Sub CanalVenda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CanalVenda, Source, X, Y)
End Sub

Private Sub CanalVenda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CanalVenda, Button, Shift, X, Y)
End Sub

Private Sub Mensagem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Mensagem, Source, X, Y)
End Sub

Private Sub Mensagem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Mensagem, Button, Shift, X, Y)
End Sub

Private Sub PesoBruto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PesoBruto, Source, X, Y)
End Sub

Private Sub PesoBruto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PesoBruto, Button, Shift, X, Y)
End Sub

Private Sub CanalVendaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CanalVendaLabel, Source, X, Y)
End Sub

Private Sub CanalVendaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CanalVendaLabel, Button, Shift, X, Y)
End Sub

Private Sub MensagemLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MensagemLabel, Source, X, Y)
End Sub

Private Sub MensagemLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MensagemLabel, Button, Shift, X, Y)
End Sub

Private Sub Placa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Placa, Source, X, Y)
End Sub

Private Sub Placa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Placa, Button, Shift, X, Y)
End Sub

Private Sub PlacaUF_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PlacaUF, Source, X, Y)
End Sub

Private Sub PlacaUF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PlacaUF, Button, Shift, X, Y)
End Sub

Private Sub Transportadora_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Transportadora, Source, X, Y)
End Sub

Private Sub Transportadora_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Transportadora, Button, Shift, X, Y)
End Sub

Private Sub FilialEntrega_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEntrega, Source, X, Y)
End Sub

Private Sub FilialEntrega_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEntrega, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub TotalPercentualComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualComissao, Source, X, Y)
End Sub

Private Sub TotalPercentualComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualComissao, Button, Shift, X, Y)
End Sub

Private Sub TotalValorComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorComissao, Source, X, Y)
End Sub

Private Sub TotalValorComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorComissao, Button, Shift, X, Y)
End Sub

'Private Sub LabelTotaisComissoes_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelTotaisComissoes, Source, X, Y)
'End Sub

'Private Sub LabelTotaisComissoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelTotaisComissoes, Button, Shift, X, Y)
'End Sub

Private Sub StatusPedido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(StatusPedido, Source, X, Y)
End Sub

Private Sub StatusPedido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(StatusPedido, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub FilialPedido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialPedido, Source, X, Y)
End Sub

Private Sub FilialPedido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialPedido, Button, Shift, X, Y)
End Sub

Private Sub Cliente_Change()
    Call objCT.Cliente_Change
End Sub

'################################
'Inserido por Wagner
Private Sub BotaoDesfazer_Click()
    Call objCT.BotaoDesfazer_Click
End Sub
'################################

Private Sub BotaoProjetos_Click()
    Call objCT.BotaoProjetos_Click
End Sub


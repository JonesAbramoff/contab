VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl NFiscalRemPedido 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4665
      Index           =   3
      Left            =   60
      TabIndex        =   168
      Top             =   1065
      Visible         =   0   'False
      Width           =   9345
      Begin VB.Frame FrameNFiscalOriginal 
         Caption         =   "Nota Fiscal Original"
         Height          =   495
         Left            =   105
         TabIndex        =   223
         Top             =   4140
         Width           =   9045
         Begin VB.ComboBox SerieNFiscalOriginal 
            Height          =   315
            Left            =   1815
            TabIndex        =   224
            Top             =   135
            Width           =   765
         End
         Begin MSMask.MaskEdBox NFiscalOriginal 
            Height          =   300
            Left            =   5310
            TabIndex        =   225
            Top             =   150
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label SerieNFOriginalLabel 
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
            Height          =   195
            Left            =   1215
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   227
            Top             =   180
            Width           =   510
         End
         Begin VB.Label NFiscalOriginalLabel 
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
            Left            =   4530
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   226
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Complemento"
         Height          =   2415
         Index           =   5
         Left            =   105
         TabIndex        =   182
         Top             =   1740
         Width           =   9030
         Begin VB.TextBox Mensagem 
            Height          =   765
            Left            =   1920
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   48
            Top             =   1230
            Width           =   7005
         End
         Begin VB.TextBox MensagemCorpo 
            Height          =   765
            Left            =   1920
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   435
            Width           =   7005
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
            Left            =   1920
            TabIndex        =   46
            Top             =   150
            Value           =   1  'Checked
            Width           =   4755
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   315
            Left            =   5400
            TabIndex        =   50
            Top             =   2040
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   315
            Left            =   1920
            TabIndex        =   49
            Top             =   2040
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumPedidoTerc 
            Height          =   315
            Left            =   7875
            TabIndex        =   51
            Top             =   2040
            Visible         =   0   'False
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
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
            Height          =   195
            Index           =   903
            Left            =   840
            TabIndex        =   187
            Top             =   2070
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
            Index           =   907
            Left            =   4125
            TabIndex        =   186
            Top             =   2100
            Width           =   1200
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   185
            Top             =   1275
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
            Index           =   52
            Left            =   150
            TabIndex        =   184
            Top             =   480
            Width           =   1665
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
            Index           =   15
            Left            =   7215
            TabIndex        =   183
            Top             =   2100
            Visible         =   0   'False
            Width           =   1305
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Volumes"
         Height          =   510
         Index           =   3
         Left            =   105
         TabIndex        =   177
         Top             =   1230
         Width           =   9030
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7140
            MaxLength       =   20
            TabIndex        =   45
            Top             =   135
            Width           =   1440
         End
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   5280
            TabIndex        =   44
            Top             =   135
            Width           =   1335
         End
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   3120
            TabIndex        =   43
            Top             =   135
            Width           =   1335
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1395
            TabIndex        =   42
            Top             =   135
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
            Index           =   906
            Left            =   6750
            TabIndex        =   181
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
            Height          =   195
            Index           =   908
            Left            =   300
            TabIndex        =   180
            Top             =   180
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
            Index           =   909
            Left            =   2295
            TabIndex        =   179
            Top             =   195
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
            Index           =   910
            Left            =   4560
            TabIndex        =   178
            Top             =   195
            Width           =   600
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dados de Entrega"
         Height          =   1260
         Index           =   6
         Left            =   105
         TabIndex        =   169
         Top             =   -15
         Width           =   9030
         Begin VB.Frame Frame6 
            Caption         =   "Frete por conta"
            Height          =   735
            Index           =   1
            Left            =   240
            TabIndex        =   172
            Top             =   480
            Width           =   1620
            Begin VB.ComboBox TipoFrete 
               Height          =   315
               Left            =   90
               TabIndex        =   188
               Top             =   300
               Width           =   1440
            End
         End
         Begin VB.ComboBox FilialEntrega 
            Height          =   315
            Left            =   1935
            TabIndex        =   36
            Top             =   150
            Width           =   2655
         End
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   6360
            TabIndex        =   37
            Top             =   150
            Width           =   2235
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   3300
            MaxLength       =   10
            TabIndex        =   38
            Top             =   540
            Width           =   1290
         End
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   3285
            TabIndex        =   39
            Top             =   885
            Width           =   735
         End
         Begin VB.Frame Frame10 
            Caption         =   "Redespacho"
            Height          =   750
            Index           =   20
            Left            =   4800
            TabIndex        =   170
            Top             =   465
            Width           =   3900
            Begin VB.ComboBox TranspRedespacho 
               Height          =   315
               Left            =   1560
               TabIndex        =   40
               Top             =   135
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
               Left            =   1560
               TabIndex        =   41
               Top             =   450
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
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   171
               Top             =   180
               Width           =   1365
            End
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   900
            Left            =   690
            TabIndex        =   176
            Top             =   195
            Width           =   1185
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
            Left            =   4965
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   175
            Top             =   180
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
            Index           =   912
            Left            =   1980
            TabIndex        =   174
            Top             =   600
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
            Index           =   911
            Left            =   1995
            TabIndex        =   173
            Top             =   915
            Width           =   1245
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4665
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   9315
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
         Left            =   7470
         TabIndex        =   34
         Top             =   4245
         Width           =   1830
      End
      Begin VB.Frame Frame2 
         Caption         =   "Totais"
         Height          =   1290
         Index           =   1
         Left            =   30
         TabIndex        =   190
         Top             =   2925
         Width           =   9285
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   90
            TabIndex        =   25
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
            TabIndex        =   191
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   28
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
            TabIndex        =   29
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
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   215
            Top             =   405
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label ValorProdutos2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   214
            Top             =   405
            Width           =   1140
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
            TabIndex        =   213
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
            Index           =   8
            Left            =   6735
            TabIndex        =   212
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
            Index           =   4
            Left            =   2790
            TabIndex        =   211
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
            Index           =   3
            Left            =   1470
            TabIndex        =   210
            Top             =   705
            Width           =   1125
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
            Index           =   5
            Left            =   105
            TabIndex        =   209
            Top             =   705
            Width           =   1125
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
            Index           =   0
            Left            =   5430
            TabIndex        =   208
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label ISSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5400
            TabIndex        =   207
            Top             =   405
            Width           =   1140
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
            Index           =   16
            Left            =   5430
            TabIndex        =   206
            Top             =   705
            Width           =   1125
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
            Index           =   6
            Left            =   4125
            TabIndex        =   205
            Top             =   705
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
            TabIndex        =   204
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
            TabIndex        =   203
            Top             =   210
            Width           =   1065
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
            TabIndex        =   202
            Top             =   210
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
            TabIndex        =   201
            Top             =   210
            Width           =   1170
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
            TabIndex        =   200
            Top             =   195
            Width           =   1065
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
            TabIndex        =   199
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   198
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4065
            TabIndex        =   197
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2745
            TabIndex        =   196
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1410
            TabIndex        =   195
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   90
            TabIndex        =   194
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   193
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8055
            TabIndex        =   192
            Top             =   915
            Width           =   1140
         End
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
         Left            =   1833
         TabIndex        =   31
         Top             =   4245
         Width           =   1620
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
         Left            =   15
         TabIndex        =   30
         Top             =   4245
         Visible         =   0   'False
         Width           =   1620
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
         Left            =   5469
         TabIndex        =   33
         Top             =   4245
         Width           =   1800
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
         Left            =   3651
         TabIndex        =   32
         Top             =   4245
         Width           =   1620
      End
      Begin VB.Frame Frame6 
         Caption         =   "Itens"
         Height          =   2955
         Index           =   7
         Left            =   30
         TabIndex        =   88
         Top             =   -15
         Width           =   9285
         Begin VB.ComboBox TabPrecoItemNF 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "NFiscalRemPedidoArtmill.ctx":0000
            Left            =   4620
            List            =   "NFiscalRemPedidoArtmill.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   222
            Top             =   1215
            Width           =   1800
         End
         Begin MSMask.MaskEdBox ComissaoItemNF 
            Height          =   255
            Left            =   6960
            TabIndex        =   221
            Top             =   1650
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
         Begin MSMask.MaskEdBox PrecoTotalB 
            Height          =   225
            Left            =   7020
            TabIndex        =   189
            Top             =   1275
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
         Begin MSMask.MaskEdBox PrecoDesc 
            Height          =   225
            Left            =   5415
            TabIndex        =   162
            Top             =   1575
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
         Begin VB.TextBox DescricaoItem2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4050
            MaxLength       =   50
            TabIndex        =   161
            Top             =   1245
            Width           =   1305
         End
         Begin VB.TextBox DescricaoItem3 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5475
            MaxLength       =   50
            TabIndex        =   160
            Top             =   1230
            Width           =   1305
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1080
            MaxLength       =   250
            TabIndex        =   23
            Top             =   765
            Width           =   2385
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1785
            TabIndex        =   17
            Text            =   "UnidadeMed"
            Top             =   285
            Width           =   855
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   7425
            TabIndex        =   22
            Top             =   390
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
            Left            =   6150
            TabIndex        =   21
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   3780
            TabIndex        =   19
            Top             =   345
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2670
            TabIndex        =   18
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   270
            TabIndex        =   16
            Top             =   285
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   4995
            TabIndex        =   20
            Top             =   330
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
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   4170
            TabIndex        =   129
            Top             =   1875
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1800
            Left            =   60
            TabIndex        =   24
            Top             =   180
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   3175
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
      Height          =   4665
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Top             =   1065
      Width           =   9345
      Begin VB.Frame FrameFornEntTerc 
         Caption         =   "Dados do Fornecedor responsável pela armazenagem e entrega dos produtos"
         Enabled         =   0   'False
         Height          =   630
         Left            =   315
         TabIndex        =   216
         Top             =   2730
         Width           =   8730
         Begin VB.ComboBox FilialFornEntTerc 
            Height          =   315
            Left            =   5805
            TabIndex        =   9
            Top             =   210
            Width           =   1860
         End
         Begin MSMask.MaskEdBox FornEntTerc 
            Height          =   315
            Left            =   2295
            TabIndex        =   8
            Top             =   210
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label FornEntTercLabel 
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
            Left            =   1185
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   218
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label LabelFilial 
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
            Left            =   5295
            TabIndex        =   217
            Top             =   270
            Width           =   465
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Outros"
         Height          =   600
         Left            =   345
         TabIndex        =   92
         Top             =   4005
         Width           =   8745
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   2310
            TabIndex        =   14
            Top             =   195
            Width           =   1485
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
            Index           =   39
            Left            =   4500
            TabIndex        =   164
            Top             =   255
            Width           =   420
         End
         Begin VB.Label NumNFe 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4995
            TabIndex        =   163
            Top             =   210
            Width           =   3240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tabela de Preços:"
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
            Left            =   660
            TabIndex        =   93
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   2070
         Index           =   0
         Left            =   315
         TabIndex        =   83
         Top             =   30
         Width           =   8730
         Begin VB.CheckBox EletronicaFed 
            Caption         =   "Eletrônica Federal"
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
            Left            =   5850
            TabIndex        =   4
            Top             =   735
            Width           =   1935
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
            Left            =   4260
            TabIndex        =   10
            Top             =   1650
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   5820
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1620
            Width           =   2550
         End
         Begin VB.ComboBox TipoNFiscal 
            Height          =   315
            ItemData        =   "NFiscalRemPedidoArtmill.ctx":0004
            Left            =   2340
            List            =   "NFiscalRemPedidoArtmill.ctx":0006
            TabIndex        =   144
            Top             =   210
            Width           =   3210
         End
         Begin VB.CommandButton BotaoLimparNF 
            Height          =   300
            Left            =   5175
            Picture         =   "NFiscalRemPedidoArtmill.ctx":0008
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpar o Número"
            Top             =   690
            Width           =   345
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   2325
            TabIndex        =   3
            Top             =   690
            Width           =   765
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   2310
            TabIndex        =   6
            Top             =   1650
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
            Left            =   600
            TabIndex        =   167
            Top             =   1245
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
            Left            =   2325
            TabIndex        =   166
            Top             =   1185
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
            Left            =   1575
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   158
            Top             =   1695
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
            Left            =   5220
            TabIndex        =   157
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label8 
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
            Index           =   1
            Left            =   1800
            TabIndex        =   145
            Top             =   270
            Width           =   450
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
            Left            =   5205
            TabIndex        =   94
            Top             =   1260
            Width           =   615
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5865
            TabIndex        =   95
            Top             =   1200
            Width           =   1560
         End
         Begin VB.Label NFiscal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3945
            TabIndex        =   96
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label NFiscalLabel 
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
            Left            =   3180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   97
            Top             =   750
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
            Left            =   1770
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   98
            Top             =   765
            Width           =   510
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente/Fornecedor"
         Height          =   630
         Index           =   2
         Left            =   315
         TabIndex        =   84
         Top             =   2115
         Width           =   8730
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
            Height          =   195
            Left            =   1320
            TabIndex        =   101
            Top             =   285
            Width           =   660
         End
         Begin VB.Label Cliente 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2040
            TabIndex        =   100
            Top             =   240
            Width           =   2250
         End
         Begin VB.Label Fornecedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2040
            TabIndex        =   154
            Top             =   240
            Width           =   2265
         End
         Begin VB.Label FornecedorLabel 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   930
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   146
            Top             =   285
            Width           =   1035
         End
         Begin VB.Label Filial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5850
            TabIndex        =   99
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label3 
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
            Left            =   5325
            TabIndex        =   102
            Top             =   285
            Width           =   555
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datas"
         Height          =   630
         Index           =   4
         Left            =   330
         TabIndex        =   85
         Top             =   3345
         Width           =   8730
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   2415
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1365
            TabIndex        =   11
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownSaida 
            Height          =   300
            Left            =   5415
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataSaida 
            Height          =   300
            Left            =   4365
            TabIndex        =   12
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HoraSaida 
            Height          =   300
            Left            =   7395
            TabIndex        =   13
            Top             =   240
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
            Caption         =   "Hora Saída:"
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
            Left            =   6270
            TabIndex        =   143
            Top             =   285
            Width           =   1050
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Data Saída:"
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
            Left            =   3225
            TabIndex        =   103
            Top             =   285
            Width           =   1050
         End
         Begin VB.Label Label5 
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
            Left            =   525
            TabIndex        =   104
            Top             =   285
            Width           =   780
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
      Height          =   495
      Left            =   5295
      TabIndex        =   148
      Top             =   150
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4545
      Index           =   5
      Left            =   225
      TabIndex        =   60
      Top             =   1080
      Visible         =   0   'False
      Width           =   9150
      Begin TelasFATArt.TabTributacaoFat TabTrib 
         Height          =   4620
         Left            =   120
         TabIndex        =   165
         Top             =   30
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   8149
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4545
      Index           =   6
      Left            =   180
      TabIndex        =   61
      Top             =   1080
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4680
         TabIndex        =   159
         Tag             =   "1"
         Top             =   1440
         Width           =   870
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   72
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   71
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6300
         TabIndex        =   74
         Top             =   1470
         Visible         =   0   'False
         Width           =   2625
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
         Height          =   510
         Left            =   6330
         TabIndex        =   79
         Top             =   630
         Width           =   1245
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
         Height          =   375
         Left            =   6330
         TabIndex        =   77
         Top             =   120
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   7740
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   810
         Width           =   1260
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   89
         Top             =   3420
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
            TabIndex        =   105
            Top             =   675
            Visible         =   0   'False
            Width           =   1440
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
            Height          =   195
            Index           =   7
            Left            =   1125
            TabIndex        =   106
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   107
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   108
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
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
         Height          =   375
         Left            =   7710
         TabIndex        =   78
         Top             =   135
         Width           =   1245
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
         Left            =   3465
         TabIndex        =   65
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4665
         TabIndex        =   70
         Top             =   1875
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
         TabIndex        =   66
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         Left            =   1665
         TabIndex        =   90
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         Left            =   3795
         TabIndex        =   62
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
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   0
         TabIndex        =   73
         Top             =   1155
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
         Left            =   6300
         TabIndex        =   75
         Top             =   1470
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
         Left            =   6300
         TabIndex        =   76
         Top             =   1470
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
         Index           =   21
         Left            =   45
         TabIndex        =   109
         Top             =   135
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   110
         Top             =   105
         Width           =   1530
      End
      Begin VB.Label CTBLabel 
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
         Index           =   14
         Left            =   4230
         TabIndex        =   111
         Top             =   585
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   112
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   113
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label CTBLabel 
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
         Index           =   13
         Left            =   1995
         TabIndex        =   114
         Top             =   570
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
         TabIndex        =   115
         Top             =   930
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
         Left            =   6300
         TabIndex        =   116
         Top             =   1230
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
         Left            =   6330
         TabIndex        =   117
         Top             =   1230
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
         Left            =   6315
         TabIndex        =   118
         Top             =   1200
         Visible         =   0   'False
         Width           =   2490
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
         Height          =   195
         Index           =   1
         Left            =   7755
         TabIndex        =   119
         Top             =   585
         Width           =   690
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
         TabIndex        =   120
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   121
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2445
         TabIndex        =   122
         Top             =   3015
         Width           =   1155
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
         TabIndex        =   123
         Top             =   540
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
         TabIndex        =   124
         Top             =   150
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
         TabIndex        =   125
         Top             =   150
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame17"
      Height          =   4545
      Index           =   7
      Left            =   195
      TabIndex        =   130
      Top             =   1065
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CommandButton BotaoSerie 
         Caption         =   "Séries"
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
         Left            =   120
         TabIndex        =   155
         Top             =   4170
         Width           =   1665
      End
      Begin VB.Frame Frame18 
         Caption         =   "Rastreamento do Produto"
         Height          =   4050
         Left            =   60
         TabIndex        =   132
         Top             =   30
         Width           =   9030
         Begin VB.ComboBox ProdutoRastro 
            Height          =   315
            ItemData        =   "NFiscalRemPedidoArtmill.ctx":053A
            Left            =   885
            List            =   "NFiscalRemPedidoArtmill.ctx":0547
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   960
            Width           =   1740
         End
         Begin VB.ComboBox EscaninhoRastro 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "NFiscalRemPedidoArtmill.ctx":0566
            Left            =   3705
            List            =   "NFiscalRemPedidoArtmill.ctx":0573
            Style           =   2  'Dropdown List
            TabIndex        =   133
            Top             =   270
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSMask.MaskEdBox UMRastro 
            Height          =   240
            Left            =   3075
            TabIndex        =   134
            Top             =   270
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
            Left            =   135
            TabIndex        =   135
            Top             =   840
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
            Left            =   1650
            TabIndex        =   136
            Top             =   300
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
            Left            =   1845
            TabIndex        =   137
            Top             =   480
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
            Left            =   2805
            TabIndex        =   138
            Top             =   480
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteDataRastro 
            Height          =   255
            Left            =   5580
            TabIndex        =   139
            Top             =   480
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
            Left            =   3960
            TabIndex        =   140
            Top             =   465
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
            Left            =   6735
            TabIndex        =   141
            Top             =   495
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
            TabIndex        =   142
            Top             =   255
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
         TabIndex        =   131
         Top             =   4170
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4545
      Index           =   4
      Left            =   210
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CheckBox ImprimeRomaneio 
         Caption         =   "Imprime Romaneio na Gravação."
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
         Left            =   480
         TabIndex        =   128
         Top             =   120
         Width           =   3270
      End
      Begin VB.CommandButton BotaoLocalizacao 
         Caption         =   "Localização de Produto"
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
         Left            =   5970
         TabIndex        =   59
         Top             =   4020
         Width           =   2460
      End
      Begin VB.Frame Frame6 
         Caption         =   "Localização dos Produtos"
         Height          =   3390
         Index           =   4
         Left            =   420
         TabIndex        =   91
         Top             =   510
         Width           =   8100
         Begin MSMask.MaskEdBox UnidadeMedEst 
            Height          =   225
            Left            =   6345
            TabIndex        =   57
            Top             =   990
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
         Begin MSMask.MaskEdBox ProdutoAlmox 
            Height          =   225
            Left            =   1860
            TabIndex        =   53
            Top             =   690
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
         Begin MSMask.MaskEdBox Almox 
            Height          =   225
            Left            =   3120
            TabIndex        =   54
            Top             =   675
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
         Begin MSMask.MaskEdBox QuantAlocada 
            Height          =   225
            Left            =   4545
            TabIndex        =   55
            Top             =   705
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
         Begin MSMask.MaskEdBox ItemNFiscal 
            Height          =   225
            Left            =   1140
            TabIndex        =   52
            Top             =   660
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
         Begin MSMask.MaskEdBox QuantVendida 
            Height          =   225
            Left            =   5760
            TabIndex        =   56
            Top             =   690
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
         Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
            Height          =   1980
            Left            =   255
            TabIndex        =   58
            Top             =   360
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   3493
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
   Begin VB.PictureBox Picture3 
      Height          =   525
      Left            =   6975
      ScaleHeight     =   465
      ScaleWidth      =   2385
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   150
      Width           =   2445
      Begin VB.CommandButton BotaoImprimirPI 
         Height          =   345
         Left            =   1890
         Picture         =   "NFiscalRemPedidoArtmill.ctx":0592
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Imprimir Pedido Interno"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   75
         Picture         =   "NFiscalRemPedidoArtmill.ctx":0694
         Style           =   1  'Graphical
         TabIndex        =   149
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   985
         Picture         =   "NFiscalRemPedidoArtmill.ctx":07EE
         Style           =   1  'Graphical
         TabIndex        =   151
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   1440
         Picture         =   "NFiscalRemPedidoArtmill.ctx":0D20
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   530
         Picture         =   "NFiscalRemPedidoArtmill.ctx":0E9E
         Style           =   1  'Graphical
         TabIndex        =   150
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   390
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Pedido"
      Height          =   645
      Index           =   0
      Left            =   45
      TabIndex        =   81
      Top             =   30
      Width           =   4725
      Begin VB.CommandButton BotaoPedido 
         Height          =   285
         Left            =   3840
         Picture         =   "NFiscalRemPedidoArtmill.ctx":1028
         Style           =   1  'Graphical
         TabIndex        =   220
         ToolTipText     =   "Trazer Dados"
         Top             =   240
         Width           =   360
      End
      Begin VB.CommandButton BotaoVerPedido 
         Height          =   285
         Left            =   4260
         Picture         =   "NFiscalRemPedidoArtmill.ctx":13FA
         Style           =   1  'Graphical
         TabIndex        =   219
         ToolTipText     =   "Trazer a tela de cadastro"
         Top             =   240
         Width           =   360
      End
      Begin VB.ComboBox FilialPedido 
         Height          =   315
         Left            =   2295
         TabIndex        =   1
         Top             =   240
         Width           =   1500
      End
      Begin MSMask.MaskEdBox Pedido 
         Height          =   300
         Left            =   885
         TabIndex        =   0
         Top             =   255
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
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
         Index           =   2
         Left            =   1830
         TabIndex        =   126
         Top             =   300
         Width           =   465
      End
      Begin VB.Label PedidoLabel 
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
         Height          =   255
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   127
         Top             =   300
         Width           =   720
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5040
      Left            =   15
      TabIndex        =   82
      Top             =   735
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   8890
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
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
            Caption         =   "Almoxarifado"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "NFiscalRemPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private objCT As CTNFiscalPedido
Private WithEvents objCT1 As CTNFiscalRemPedido
Attribute objCT1.VB_VarHelpID = -1
Private objCT2 As CTNFiscalRem

Private Sub Botao_Click(Index As Integer)
    Call objCT1.BotaoPedido_Click
End Sub

Private Sub BotaoGrade_Click()
    Call objCT.BotaoGrade_Click
End Sub

Private Sub BotaoExcluir_Click()
    Call objCT1.BotaoExcluir_Click
End Sub

'Rastreamento
Private Sub BotaoLotes_Click()
    Call objCT1.BotaoLotes_Click
End Sub

Private Sub objCT1_Unload()
   RaiseEvent Unload
End Sub

Public Sub TipoNFiscal_Click()
    Call objCT1.TipoNFiscal_Click
End Sub

Public Sub TipoNFiscal_Change()
    Call objCT1.TipoNFiscal_Change
End Sub

Public Sub TipoNFiscal_Validate(Cancel As Boolean)
    Call objCT1.TipoNFiscal_Validate(Cancel)
End Sub

Private Sub ItemNFRastro_Change()
     Call objCT1.ItemNFRastro_Change
End Sub

Private Sub ItemNFRastro_GotFocus()
     Call objCT1.ItemNFRastro_GotFocus
End Sub

Private Sub ItemNFRastro_KeyPress(KeyAscii As Integer)
     Call objCT1.ItemNFRastro_KeyPress(KeyAscii)
End Sub

Private Sub ItemNFRastro_Validate(Cancel As Boolean)
     Call objCT1.ItemNFRastro_Validate(Cancel)
End Sub

Private Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT1.GridRastro_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridRastro_Click()
     Call objCT1.GridRastro_Click
End Sub

Private Sub GridRastro_EnterCell()
     Call objCT1.GridRastro_EnterCell
End Sub

Private Sub GridRastro_GotFocus()
     Call objCT1.GridRastro_GotFocus
End Sub

Private Sub GridRastro_KeyPress(KeyAscii As Integer)
     Call objCT1.GridRastro_KeyPress(KeyAscii)
End Sub

Private Sub GridRastro_LeaveCell()
     Call objCT1.GridRastro_LeaveCell
End Sub

Private Sub GridRastro_Validate(Cancel As Boolean)
     Call objCT1.GridRastro_Validate(Cancel)
End Sub

Private Sub GridRastro_Scroll()
     Call objCT1.GridRastro_Scroll
End Sub

Private Sub GridRastro_RowColChange()
     Call objCT1.GridRastro_RowColChange
End Sub

Private Sub LoteRastro_Change()
     Call objCT1.LoteRastro_Change
End Sub

Private Sub LoteRastro_GotFocus()
     Call objCT1.LoteRastro_GotFocus
End Sub

Private Sub LoteRastro_KeyPress(KeyAscii As Integer)
     Call objCT1.LoteRastro_KeyPress(KeyAscii)
End Sub

Private Sub LoteRastro_Validate(Cancel As Boolean)
     Call objCT1.LoteRastro_Validate(Cancel)
End Sub

Private Sub FilialOPRastro_Change()
     Call objCT1.FilialOPRastro_Change
End Sub

Private Sub FilialOPRastro_GotFocus()
     Call objCT1.FilialOPRastro_GotFocus
End Sub

Private Sub FilialOPRastro_KeyPress(KeyAscii As Integer)
     Call objCT1.FilialOPRastro_KeyPress(KeyAscii)
End Sub

Private Sub FilialOPRastro_Validate(Cancel As Boolean)
     Call objCT1.FilialOPRastro_Validate(Cancel)
End Sub

Private Sub QuantLoteRastro_Change()
     Call objCT1.QuantLoteRastro_Change
End Sub

Private Sub QuantLoteRastro_GotFocus()
     Call objCT1.QuantLoteRastro_GotFocus
End Sub

Private Sub QuantLoteRastro_KeyPress(KeyAscii As Integer)
     Call objCT1.QuantLoteRastro_KeyPress(KeyAscii)
End Sub

Private Sub QuantLoteRastro_Validate(Cancel As Boolean)
     Call objCT1.QuantLoteRastro_Validate(Cancel)
End Sub

Private Sub AlmoxRastro_Change()
     Call objCT1.AlmoxRastro_Change
End Sub

Private Sub AlmoxRastro_GotFocus()
     Call objCT1.AlmoxRastro_GotFocus
End Sub

Private Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
     Call objCT1.AlmoxRastro_KeyPress(KeyAscii)
End Sub

Private Sub AlmoxRastro_Validate(Cancel As Boolean)
     Call objCT1.AlmoxRastro_Validate(Cancel)
End Sub

Private Sub EscaninhoRastro_Change()
     Call objCT1.QuantLoteRastro_Change
End Sub

'Private Sub EscaninhoRastro_GotFocus()
'     Call objCT2.EscaninhoRastro_GotFocus
'End Sub
'
'Private Sub EscaninhoRastro_KeyPress(KeyAscii As Integer)
'     Call objCT2.EscaninhoRastro_KeyPress(KeyAscii)
'End Sub
'
'Private Sub EscaninhoRastro_Validate(Cancel As Boolean)
'     Call objCT2.EscaninhoRastro_Validate(Cancel)
'End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTNFiscalPedido
    Set objCT.objUserControl = Me

    Set objCT1 = New CTNFiscalRemPedido
    Set objCT1.objCTNFiscalPedido = objCT

    objCT1.objCTNFiscalPedido.sName = objCT1.Name
       
    Set objCT2 = New CTNFiscalRem
    Set objCT1.objCTNFiscalRem = objCT2
    Set objCT2.objUserControl = Me
    
    Set objCT1.gobjInfoUsu = New CTNFiscalRemPedidoVGArt
    Set objCT1.gobjInfoUsu.gobjTelaUsu = New CTNFiscalRemPedidoArt

    Set objCT.gobjInfoUsu = New CTNFiscalPedidoVGArt
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTNFiscalPedidoArt

    Set objCT2.gobjInfoUsu = New CTNFiscalRemVGArt
    Set objCT2.gobjInfoUsu.gobjTelaUsu = New CTNFiscalRemArt

End Sub

Private Sub ImprimeRomaneio_Click()
    Call objCT.ImprimeRomaneio_Click
End Sub

Private Sub BotaoLimparNF_Click()
     Call objCT.BotaoLimparNF_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT1.BotaoFechar_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataSaida_GotFocus()
     Call objCT.DataSaida_GotFocus
End Sub

Private Sub TipoFrete_Click()
     Call objCT.TipoFrete_Click
End Sub

Private Sub Pedido_GotFocus()
     Call objCT.Pedido_GotFocus
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT1.BotaoLimpar_Click
End Sub

Private Sub BotaoVerPedido_Click()
     Call objCT.BotaoVerPedido_Click
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
End Sub

Private Sub DataSaida_Change()
     Call objCT.DataSaida_Change
End Sub

Private Sub HoraSaida_Change()
     Call objCT.HoraSaida_Change
End Sub

Private Sub HoraSaida_Validate(Cancel As Boolean)
     Call objCT.HoraSaida_Validate(Cancel)
End Sub

Private Sub HoraSaida_GotFocus()
     Call objCT.HoraSaida_GotFocus
End Sub

Private Sub FilialEntrega_Change()
     Call objCT.FilialEntrega_Change
End Sub

Private Sub FilialPedido_Change()
     Call objCT.FilialPedido_Change
End Sub

Private Sub FilialPedido_Click()
     Call objCT.FilialPedido_Click
End Sub

Public Sub Form_Activate()
    Call objCT1.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub Mensagem_Change()
     Call objCT.Mensagem_Change
End Sub

Private Sub NFiscal_Change()
     Call objCT.NFiscal_Change
End Sub

Private Sub Pedido_Change()
     Call objCT.Pedido_Change
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

Private Sub Serie_Change()
     Call objCT.Serie_Change
End Sub

Private Sub TabelaPreco_Change()
     Call objCT.TabelaPreco_Change
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDespesas_Change()
     Call objCT.ValorDespesas_Change
End Sub

Private Sub ValorProdutos_Change()
     Call objCT.ValorProdutos_Change
End Sub

Private Sub VolumeEspecie_Change()
     Call objCT.VolumeEspecie_Change
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Private Sub VolumeEspecie_Click()
     Call objCT.VolumeEspecie_Click
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Private Sub VolumeEspecie_Validate(Cancel As Boolean)
    Call objCT.VolumeEspecie_Validate(Cancel)
End Sub

Private Sub VolumeMarca_Change()
     Call objCT.VolumeMarca_Change
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Private Sub VolumeMarca_Click()
     Call objCT.VolumeMarca_Click
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Private Sub VolumeMarca_Validate(Cancel As Boolean)
    Call objCT.VolumeMarca_Validate(Cancel)
End Sub

Private Sub VolumeNumero_Change()
     Call objCT.VolumeNumero_Change
End Sub

Private Sub VolumeQuant_Change()
     Call objCT.VolumeQuant_Change
End Sub

Private Sub Serie_Click()
     Call objCT.Serie_Click
End Sub

Private Sub Serie_Validate(Cancel As Boolean)
     Call objCT.Serie_Validate(Cancel)
End Sub

Private Sub SerieLabel_Click()
     Call objCT.SerieLabel_Click
End Sub

Private Sub NFiscalLabel_Click()
'Janaina
     Call objCT1.NFiscalLabel_Click
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

Private Sub DataSaida_Validate(Cancel As Boolean)
     Call objCT.DataSaida_Validate(Cancel)
End Sub

Private Sub UpDownSaida_DownClick()
     Call objCT.UpDownSaida_DownClick
End Sub

Private Sub UpDownSaida_UpClick()
     Call objCT.UpDownSaida_UpClick
End Sub

Private Sub TabelaPreco_Click()
     Call objCT.TabelaPreco_Click
End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)
     Call objCT.TabelaPreco_Validate(Cancel)
End Sub

Private Sub ValorFrete_Change()
     Call objCT.ValorFrete_Change
End Sub

Private Sub ValorSeguro_Change()
     Call objCT.ValorSeguro_Change
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

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub FilialEntrega_Validate(Cancel As Boolean)
     Call objCT.FilialEntrega_Validate(Cancel)
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub PlacaUF_Validate(Cancel As Boolean)
     Call objCT.PlacaUF_Validate(Cancel)
End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)
     Call objCT.PesoBruto_Validate(Cancel)
End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)
     Call objCT.PesoLiquido_Validate(Cancel)
End Sub

Private Sub BotaoEstoqueProd_Click()
     Call objCT.BotaoEstoqueProd_Click
End Sub

Private Sub FilialPedido_Validate(Cancel As Boolean)
     Call objCT.FilialPedido_Validate(Cancel)
End Sub
'Janaina
Private Sub BotaoPedido_Click()
     Call objCT1.BotaoPedido_Click
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
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
'janaina
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
'Janaina
Private Sub GridItens_LeaveCell()
    Call objCT1.GridItens_LeaveCell 'Alterado por Leo em 17/01/02
    'agora chama a função GridItens_LeaveCell em objct1 e não mais em objct. Motivo: a função GridItens_LeaveCell chama a função ValorTotal_Calcula que em objct não chama a função de comissões.
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
    Call objCT1.GridItens_KeyDown(KeyCode, Shift) 'alterado por Leo em 17/01/02
    'agora chama a função GridItens_KeyDown em objct1 e não mais em objct. Motivo: a função GridItens_KeyDown chama a função ValorTotal_Calcula que em objct não chama a função de comissões.
End Sub

Private Sub BotaoGravar_Click()
     Call objCT1.BotaoGravar_Click
End Sub

Private Sub PedidoLabel_Click()
'Janaina
     Call objCT1.PedidoLabel_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub BotaoLocalizacao_Click()
     Call objCT1.BotaoLocalizacao_Click 'alterado por Leo em 15/01/02
End Sub

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long
    Trata_Parametros = objCT1.Trata_Parametros(objNFiscal)
End Function

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

Private Sub VolumeQuant_GotFocus()
     Call objCT.VolumeQuant_GotFocus
End Sub

Private Sub MensagemLabel_Click()
     Call objCT.MensagemLabel_Click
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT1.Form_Load_Ocx

    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)

    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel Then Exit Sub
        
        Set objCT.objUserControl = Nothing
        Set objCT = Nothing
    End If
    
    If Not (objCT1 Is Nothing) Then
        Call objCT1.Form_Unload(Cancel)
        Set objCT1 = Nothing
    End If
    
    If Not (objCT2 Is Nothing) Then
        Set objCT2 = Nothing
    End If
    
End Sub

Public Function Name() As String
    Name = objCT1.Name
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


'Private Sub Label17_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label17(Index), Source, X, Y)
'End Sub
'
'Private Sub Label17_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label17(Index), Button, Shift, X, Y)
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
'Private Sub Label13_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label13(Index), Source, X, Y)
'End Sub
'
'Private Sub Label13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label13(Index), Button, Shift, X, Y)
'End Sub
'
'Private Sub Label18_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label18(Index), Source, X, Y)
'End Sub
'
'Private Sub Label18_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label18(Index), Button, Shift, X, Y)
'End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

'Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label15(Index), Source, X, Y)
'End Sub
'
'Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
'End Sub

Private Sub Label7_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label7(Index), Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7(Index), Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label20(Index), Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20(Index), Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
End Sub

'Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label10, Source, X, Y)
'End Sub
'
'Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
'End Sub

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

'Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelTotais, Source, X, Y)
'End Sub
'
'Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
'End Sub

'Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label9, Source, X, Y)
'End Sub
'
'Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
'End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

'Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label14, Source, X, Y)
'End Sub
'
'Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
'End Sub

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

Private Sub Filial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Filial, Source, X, Y)
End Sub

Private Sub Filial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Filial, Button, Shift, X, Y)
End Sub

Private Sub Cliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cliente, Source, X, Y)
End Sub

Private Sub Cliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cliente, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
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

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
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

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
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

'colocar o label de tipo e outros...

Private Sub PedidoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PedidoLabel, Source, X, Y)
End Sub

Private Sub PedidoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PedidoLabel, Button, Shift, X, Y)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub Fornecedor_Change()
     Call objCT1.Fornecedor_Change
End Sub
'Janaina
Private Sub Cliente_Change()
     Call objCT1.Cliente_Change
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub
'Janaina
Private Sub Filial_Change()
     Call objCT2.Filial_Change
End Sub
'Janaina
Private Sub RedespachoCli_Click()
    Call objCT.RedespachoCli_Click
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

Private Sub TranspRedLabel_Click()
     Call objCT.TranspRedLabel_Click
End Sub

'#####################################
'Inserido por Wagner 15/03/2006
Private Sub BotaoSerie_Click()
    Call objCT.BotaoSerie_Click
End Sub
'#####################################

'#####################################
'Inserido por Wagner 05/06/2006
Private Sub MensagemCorpo_Change()
     Call objCT.MensagemCorpo_Change
End Sub
'#####################################

'#####################################
'Inserido por Wagner 03/08/2006
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
'#####################################

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


'##########################################
'Inserido por Wagner
Private Sub BotaoImprimirPI_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoImprimirPI_Click(objCT)
End Sub
'##########################################

Private Sub MsgAutomatica_Click()
     Call objCT.MsgAutomatica_Click
End Sub

Private Sub EletronicaFed_Click()
    Call objCT.EletronicaFed_Click
End Sub

Private Sub BotaoInfoAdic_Click()
     Call objCT.BotaoInfoAdic_Click
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


Private Sub FornEntTercLabel_Click()
     Call objCT1.FornEntTercLabel_Click
End Sub

Private Sub FornEntTerc_Change()
     Call objCT1.FornEntTerc_Change
End Sub

Private Sub FornEntTerc_Validate(Cancel As Boolean)
     Call objCT1.FornEntTerc_Validate(Cancel)
End Sub

Private Sub FilialFornEntTerc_Change()
     Call objCT1.FilialFornEntTerc_Change
End Sub

Private Sub FilialFornEntTerc_Click()
     Call objCT1.FilialFornEntTerc_Click
End Sub

Private Sub FilialFornEntTerc_Validate(Cancel As Boolean)
     Call objCT1.FilialFornEntTerc_Validate(Cancel)
End Sub

Private Sub BotaoInfoAdicItem_Click()
    Call objCT.BotaoInfoAdicItem_Click
End Sub

Private Sub ComissaoItemNF_Change()
     Call objCT.ComissaoItemNF_Change
End Sub

Private Sub ComissaoItemNF_GotFocus()
     Call objCT.ComissaoItemNF_GotFocus
End Sub

Private Sub ComissaoItemNF_KeyPress(KeyAscii As Integer)
     Call objCT.ComissaoItemNF_KeyPress(KeyAscii)
End Sub

Private Sub ComissaoItemNF_Validate(Cancel As Boolean)
     Call objCT.ComissaoItemNF_Validate(Cancel)
End Sub

Private Sub TabPrecoItemNF_Change()
     Call objCT.TabPrecoItemNF_Change
End Sub

Private Sub TabPrecoItemNF_Click()
     Call objCT.TabPrecoItemNF_Click
End Sub

Private Sub TabPrecoItemNF_GotFocus()
     Call objCT.TabPrecoItemNF_GotFocus
End Sub

Private Sub TabPrecoItemNF_KeyPress(KeyAscii As Integer)
     Call objCT.TabPrecoItemNF_KeyPress(KeyAscii)
End Sub

Private Sub TabPrecoItemNF_Validate(Cancel As Boolean)
     Call objCT.TabPrecoItemNF_Validate(Cancel)
End Sub

Private Sub NFiscalOriginal_Change()
     Call objCT2.NFiscalOriginal_Change
End Sub

Private Sub NFiscalOriginalLabel_Click()
     Call objCT1.NFiscalOriginalLabel_Click
End Sub

Private Sub NFiscalOriginal_GotFocus()
     Call objCT2.NFiscalOriginal_GotFocus
End Sub

Private Sub SerieNFOriginalLabel_Click()
     Call objCT2.SerieNFOriginalLabel_Click
End Sub

Private Sub SerieNFiscalOriginal_Change()
     Call objCT2.SerieNFiscalOriginal_Change
End Sub

Private Sub SerieNFiscalOriginal_Click()
     Call objCT2.SerieNFiscalOriginal_Click
End Sub

Private Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)
     Call objCT2.SerieNFiscalOriginal_Validate(Cancel)
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl NFiscalGR 
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4725
      Index           =   4
      Left            =   150
      TabIndex        =   26
      Top             =   1410
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame Frame10 
         Caption         =   "Dados de Entrega"
         Height          =   1080
         Index           =   13
         Left            =   120
         TabIndex        =   115
         Top             =   60
         Width           =   9030
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   6675
            TabIndex        =   32
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   3180
            MaxLength       =   10
            TabIndex        =   31
            Top             =   675
            Width           =   1290
         End
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   6690
            TabIndex        =   28
            Top             =   270
            Width           =   2235
         End
         Begin VB.ComboBox FilialEntrega 
            Height          =   315
            Left            =   3195
            TabIndex        =   27
            Top             =   285
            Width           =   1935
         End
         Begin VB.Frame Frame10 
            Caption         =   "Frete por conta"
            Height          =   795
            Index           =   15
            Left            =   165
            TabIndex        =   117
            Top             =   210
            Width           =   1605
            Begin VB.OptionButton Emitente 
               Caption         =   "Emitente"
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
               Left            =   75
               TabIndex        =   29
               Top             =   210
               Value           =   -1  'True
               Width           =   1470
            End
            Begin VB.OptionButton Destinatario 
               Caption         =   "Destinat�rio"
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
               Left            =   75
               TabIndex        =   30
               Top             =   495
               Width           =   1500
            End
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
            Left            =   5250
            TabIndex        =   161
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label30 
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
            Index           =   4
            Left            =   1860
            TabIndex        =   162
            Top             =   735
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
            Left            =   5250
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   163
            Top             =   330
            Width           =   1365
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
            Index           =   12
            Left            =   1935
            TabIndex        =   164
            Top             =   345
            Width           =   1185
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Volumes"
         Height          =   570
         Index           =   12
         Left            =   105
         TabIndex        =   118
         Top             =   1110
         Width           =   9030
         Begin VB.TextBox VolumeMarca 
            Height          =   300
            Left            =   5355
            MaxLength       =   20
            TabIndex        =   35
            Top             =   180
            Width           =   1020
         End
         Begin VB.TextBox VolumeEspecie 
            Height          =   300
            Left            =   3090
            MaxLength       =   20
            TabIndex        =   34
            Top             =   180
            Width           =   1335
         End
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7140
            MaxLength       =   20
            TabIndex        =   36
            Top             =   180
            Width           =   1440
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1395
            TabIndex        =   33
            Top             =   180
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
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
            TabIndex        =   165
            Top             =   225
            Width           =   600
         End
         Begin VB.Label Label30 
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
            Height          =   195
            Index           =   1
            Left            =   2295
            TabIndex        =   166
            Top             =   225
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
            TabIndex        =   167
            Top             =   225
            Width           =   1050
         End
         Begin VB.Label Label20 
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
            Height          =   195
            Index           =   8
            Left            =   6750
            TabIndex        =   168
            Top             =   225
            Width           =   345
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Complemento"
         Height          =   2205
         Index           =   11
         Left            =   105
         TabIndex        =   119
         Top             =   1725
         Width           =   9030
         Begin VB.TextBox Destino 
            Height          =   300
            Left            =   6105
            MaxLength       =   250
            TabIndex        =   246
            Top             =   1425
            Width           =   2820
         End
         Begin VB.TextBox Origem 
            Height          =   300
            Left            =   2055
            MaxLength       =   250
            TabIndex        =   244
            Top             =   1425
            Width           =   2820
         End
         Begin VB.ComboBox CanalVenda 
            Height          =   315
            Left            =   2070
            TabIndex        =   40
            Top             =   1020
            Width           =   1620
         End
         Begin VB.TextBox Mensagem 
            Height          =   300
            Left            =   2085
            MaxLength       =   250
            TabIndex        =   37
            Top             =   255
            Width           =   6510
         End
         Begin MSMask.MaskEdBox NumPedidoTerc 
            Height          =   300
            Left            =   6105
            TabIndex        =   41
            Top             =   1035
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   300
            Left            =   6105
            TabIndex        =   38
            Top             =   645
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   300
            Left            =   2070
            TabIndex        =   39
            Top             =   630
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorContainer 
            Height          =   300
            Left            =   2070
            TabIndex        =   248
            Top             =   1830
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorMercadoria 
            Height          =   300
            Left            =   6105
            TabIndex        =   250
            Top             =   1815
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor da Mercadoria:"
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
            Index           =   180
            Left            =   4245
            TabIndex        =   251
            Top             =   1860
            Width           =   1785
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Container:"
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
            Index           =   178
            Left            =   360
            TabIndex        =   249
            Top             =   1860
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
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
            Index           =   179
            Left            =   5310
            TabIndex        =   247
            Top             =   1455
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   177
            Left            =   1350
            TabIndex        =   245
            Top             =   1455
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Peso L�quido:"
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
            Left            =   4815
            TabIndex        =   169
            Top             =   690
            Width           =   1200
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
            Index           =   4
            Left            =   1005
            TabIndex        =   170
            Top             =   660
            Width           =   1005
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
            Index           =   5
            Left            =   4725
            TabIndex        =   171
            Top             =   1080
            Width           =   1305
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   172
            Top             =   270
            Width           =   1725
         End
         Begin VB.Label Label1 
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
            Index           =   14
            Left            =   585
            TabIndex        =   173
            Top             =   1065
            Width           =   1425
         End
      End
      Begin VB.Frame FrameNFiscalOriginal 
         Caption         =   "Nota Fiscal Original"
         Height          =   630
         Left            =   105
         TabIndex        =   120
         Top             =   4005
         Width           =   9030
         Begin VB.ComboBox SerieNFiscalOriginal 
            Height          =   315
            Left            =   2085
            TabIndex        =   42
            Top             =   210
            Width           =   765
         End
         Begin MSMask.MaskEdBox NFiscalOriginal 
            Height          =   300
            Left            =   5640
            TabIndex        =   43
            Top             =   225
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
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
            Height          =   255
            Left            =   4860
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   174
            Top             =   240
            Width           =   720
         End
         Begin VB.Label SerieNFOriginalLabel 
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
            Height          =   195
            Left            =   1485
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   175
            Top             =   255
            Width           =   510
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4725
      Index           =   3
      Left            =   150
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame Frame10 
         Caption         =   "Valores"
         Height          =   1275
         Index           =   9
         Left            =   150
         TabIndex        =   123
         Top             =   2835
         Width           =   8880
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   1755
            TabIndex        =   20
            Top             =   870
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   450
            TabIndex        =   23
            Top             =   390
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
            Left            =   4365
            TabIndex        =   22
            Top             =   870
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
            Left            =   3060
            TabIndex        =   21
            Top             =   870
            Width           =   1125
            _ExtentX        =   1984
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
            Left            =   450
            TabIndex        =   272
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   17
            Left            =   795
            TabIndex        =   271
            Top             =   675
            Width           =   315
         End
         Begin VB.Label Label17 
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
            Height          =   195
            Index           =   0
            Left            =   6105
            TabIndex        =   143
            Top             =   195
            Width           =   1005
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6045
            TabIndex        =   144
            Top             =   390
            Width           =   1125
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
            Height          =   165
            Index           =   0
            Left            =   6480
            TabIndex        =   145
            Top             =   675
            Width           =   255
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6045
            TabIndex        =   146
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label Label20 
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
            Height          =   195
            Index           =   6
            Left            =   3315
            TabIndex        =   147
            Top             =   675
            Width           =   615
         End
         Begin VB.Label Label15 
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
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   148
            Top             =   195
            Width           =   825
         End
         Begin VB.Label Label15 
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
            Height          =   195
            Index           =   2
            Left            =   2092
            TabIndex        =   149
            Top             =   675
            Width           =   450
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7365
            TabIndex        =   150
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label Label9 
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
            Height          =   195
            Left            =   4695
            TabIndex        =   151
            Top             =   675
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
            Height          =   195
            Left            =   7695
            TabIndex        =   152
            Top             =   675
            Width           =   450
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7350
            TabIndex        =   153
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label18 
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
            Height          =   195
            Index           =   0
            Left            =   7530
            TabIndex        =   154
            Top             =   195
            Width           =   765
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1755
            TabIndex        =   155
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3060
            TabIndex        =   156
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4365
            TabIndex        =   157
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
            Height          =   195
            Index           =   3
            Left            =   3390
            TabIndex        =   158
            Top             =   195
            Width           =   465
         End
         Begin VB.Label Label13 
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
            Height          =   195
            Index           =   0
            Left            =   1845
            TabIndex        =   159
            Top             =   195
            Width           =   945
         End
         Begin VB.Label Label16 
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
            Height          =   195
            Index           =   0
            Left            =   4373
            TabIndex        =   160
            Top             =   195
            Width           =   1485
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Itens"
         Height          =   2715
         Index           =   10
         Left            =   150
         TabIndex        =   122
         Top             =   60
         Width           =   8865
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   5940
            TabIndex        =   18
            Top             =   1095
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
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   5100
            TabIndex        =   15
            Top             =   645
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
            Left            =   4170
            TabIndex        =   17
            Top             =   1065
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
            Left            =   3405
            TabIndex        =   14
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2610
            TabIndex        =   13
            Top             =   420
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
         Begin MSMask.MaskEdBox UnidadeMed 
            Height          =   225
            Left            =   1980
            TabIndex        =   270
            Top             =   1035
            Width           =   510
            _ExtentX        =   900
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
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   315
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1065
            Width           =   1665
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   210
            TabIndex        =   12
            Top             =   705
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2205
            Left            =   120
            TabIndex        =   19
            Top             =   330
            Width           =   8625
            _ExtentX        =   15214
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
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
         Left            =   7410
         TabIndex        =   25
         Top             =   4245
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
         Height          =   345
         Left            =   5835
         TabIndex        =   24
         Top             =   4245
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4710
      Index           =   1
      Left            =   150
      TabIndex        =   0
      Top             =   1320
      Width           =   9195
      Begin VB.Frame Frame10 
         Caption         =   "Identifica��o"
         Height          =   1965
         Index           =   6
         Left            =   210
         TabIndex        =   111
         Top             =   120
         Width           =   8685
         Begin VB.CommandButton BotaoLimparNF 
            Height          =   300
            Left            =   5910
            Picture         =   "nfiscalgr.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpar o N�mero"
            Top             =   1545
            Width           =   345
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   1995
            TabIndex        =   3
            Top             =   1545
            Width           =   765
         End
         Begin VB.ComboBox TipoNFiscal 
            Height          =   315
            ItemData        =   "nfiscalgr.ctx":0532
            Left            =   1995
            List            =   "nfiscalgr.ctx":0534
            TabIndex        =   1
            Top             =   270
            Width           =   2835
         End
         Begin MSMask.MaskEdBox NatOpInterna 
            Height          =   300
            Left            =   1980
            TabIndex        =   2
            Top             =   888
            Width           =   552
            _ExtentX        =   979
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
            Left            =   4500
            TabIndex        =   133
            Top             =   938
            Width           =   615
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5160
            TabIndex        =   134
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label NFiscal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5160
            TabIndex        =   135
            Top             =   1545
            Width           =   735
         End
         Begin VB.Label LblNatOpInterna 
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
            Left            =   195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   136
            Top             =   945
            Width           =   1725
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
            Left            =   1395
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   137
            Top             =   1590
            Width           =   510
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
            Height          =   255
            Left            =   4410
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   138
            Top             =   1575
            Width           =   720
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
            Left            =   1455
            TabIndex        =   139
            Top             =   315
            Width           =   450
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Datas"
         Height          =   690
         Index           =   8
         Left            =   240
         TabIndex        =   252
         Top             =   2970
         Width           =   8685
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   2445
            TabIndex        =   253
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
            TabIndex        =   7
            Top             =   255
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
            Left            =   5430
            TabIndex        =   254
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataSaida 
            Height          =   300
            Left            =   4365
            TabIndex        =   8
            Top             =   255
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
            Left            =   7365
            TabIndex        =   9
            Top             =   255
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Data Sa�da:"
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
            Index           =   5
            Left            =   3225
            TabIndex        =   257
            Top             =   300
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   10
            Left            =   510
            TabIndex        =   256
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora Sa�da:"
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
            Left            =   6225
            TabIndex        =   255
            Top             =   300
            Width           =   1050
         End
      End
      Begin VB.ComboBox TabelaPreco 
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   3945
         Width           =   1875
      End
      Begin VB.Frame Frame10 
         Caption         =   "Dados do Cliente"
         Height          =   675
         Index           =   7
         Left            =   240
         TabIndex        =   116
         Top             =   2190
         Width           =   8670
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5130
            TabIndex        =   6
            Top             =   270
            Width           =   1860
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1950
            TabIndex        =   5
            Top             =   270
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   1185
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   140
            Top             =   300
            Width           =   660
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
            Index           =   11
            Left            =   4590
            TabIndex        =   141
            Top             =   300
            Width           =   465
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tabela de Pre�os:"
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
         Left            =   495
         TabIndex        =   142
         Top             =   3990
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comprovantes de Servi�os"
      Height          =   4740
      Index           =   99
      Left            =   105
      TabIndex        =   258
      Top             =   1425
      Visible         =   0   'False
      Width           =   9180
      Begin MSMask.MaskEdBox ValorContainerCon 
         Height          =   225
         Left            =   2985
         TabIndex        =   269
         Top             =   1440
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
         Left            =   3720
         TabIndex        =   268
         Top             =   945
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PrecoCon 
         Height          =   225
         Left            =   2580
         TabIndex        =   267
         Top             =   1095
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox QuantCon 
         Height          =   225
         Left            =   2595
         TabIndex        =   266
         Top             =   540
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
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
         Left            =   7005
         TabIndex        =   265
         Top             =   4080
         Width           =   1365
      End
      Begin MSMask.MaskEdBox AdValorenCon 
         Height          =   225
         Left            =   3810
         TabIndex        =   263
         Top             =   570
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ServicoCon 
         Height          =   225
         Left            =   1590
         TabIndex        =   261
         Top             =   480
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DescricaoCon 
         Height          =   225
         Left            =   1320
         TabIndex        =   264
         Top             =   930
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DataCon 
         Height          =   225
         Left            =   375
         TabIndex        =   262
         Top             =   975
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
         Left            =   300
         TabIndex        =   260
         Top             =   585
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   10
         Mask            =   "##########"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridComprovServ 
         Height          =   3450
         Left            =   225
         TabIndex        =   259
         Top             =   330
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   6085
         _Version        =   393216
         Rows            =   7
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   274
         Top             =   3930
         Width           =   1125
      End
      Begin VB.Label LabelComprovante 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   360
         TabIndex        =   273
         Top             =   4170
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4725
      Index           =   7
      Left            =   150
      TabIndex        =   64
      Top             =   1425
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame FrameTributacao 
         BorderStyle     =   0  'None
         Caption         =   "Resumo"
         Height          =   4110
         Index           =   1
         Left            =   210
         TabIndex        =   284
         Top             =   435
         Width           =   8700
         Begin VB.Frame Frame10 
            Caption         =   "IPI"
            Height          =   1620
            Index           =   18
            Left            =   4155
            TabIndex        =   335
            Top             =   900
            Width           =   2124
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
               Height          =   195
               Index           =   30
               Left            =   255
               TabIndex        =   340
               Top             =   945
               Width           =   525
            End
            Begin VB.Label IPIValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   855
               TabIndex        =   339
               Top             =   900
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   28
               Left            =   300
               TabIndex        =   338
               Top             =   465
               Width           =   495
            End
            Begin VB.Label IPIBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   870
               TabIndex        =   337
               Top             =   375
               Width           =   1080
            End
            Begin VB.Label IPICredito 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   855
               TabIndex        =   336
               Top             =   900
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cr�dito:"
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
               Height          =   195
               Index           =   36
               Left            =   990
               TabIndex        =   341
               Top             =   960
               Visible         =   0   'False
               Width           =   660
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "ICMS"
            Height          =   1635
            Index           =   4
            Left            =   435
            TabIndex        =   323
            Top             =   900
            Width           =   3600
            Begin VB.Frame Frame10 
               Caption         =   "Substituicao"
               Height          =   780
               Index           =   16
               Left            =   165
               TabIndex        =   324
               Top             =   720
               Width           =   3255
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
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
                  Left            =   1740
                  TabIndex        =   328
                  Top             =   180
                  Width           =   450
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
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
                  Left            =   390
                  TabIndex        =   327
                  Top             =   180
                  Width           =   450
               End
               Begin VB.Label ICMSSubstBase 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   375
                  TabIndex        =   326
                  Top             =   375
                  Width           =   1080
               End
               Begin VB.Label ICMSSubstValor 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   325
                  Top             =   375
                  Width           =   1080
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
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
               Index           =   24
               Left            =   1920
               TabIndex        =   333
               Top             =   165
               Width           =   630
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Base"
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
               Left            =   555
               TabIndex        =   332
               Top             =   165
               Width           =   450
            End
            Begin VB.Label ICMSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   525
               TabIndex        =   331
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label ICMSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1920
               TabIndex        =   330
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label ICMSCredito 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   1920
               TabIndex        =   329
               Top             =   405
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cr�dito"
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
               Height          =   195
               Index           =   27
               Left            =   2130
               TabIndex        =   334
               Top             =   480
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "ISS"
            Height          =   1635
            Index           =   5
            Left            =   6420
            TabIndex        =   315
            Top             =   900
            Width           =   1980
            Begin VB.CheckBox ISSIncluso 
               Caption         =   "Incluso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   630
               TabIndex        =   316
               Top             =   1350
               Width           =   1020
            End
            Begin MSMask.MaskEdBox ISSAliquota 
               Height          =   285
               Left            =   615
               TabIndex        =   317
               Top             =   630
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ISSValor 
               Height          =   285
               Left            =   615
               TabIndex        =   318
               Top             =   975
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
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
               Height          =   195
               Index           =   33
               Left            =   90
               TabIndex        =   322
               Top             =   1020
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Index           =   32
               Left            =   336
               TabIndex        =   321
               Top             =   672
               Width           =   216
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   31
               Left            =   105
               TabIndex        =   320
               Top             =   285
               Width           =   495
            End
            Begin VB.Label ISSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   615
               TabIndex        =   319
               Top             =   255
               Width           =   1110
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "CSLL"
            Height          =   570
            Index           =   3
            Left            =   4605
            TabIndex        =   312
            Top             =   3510
            Width           =   1860
            Begin MSMask.MaskEdBox CSLLRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   313
               Top             =   195
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   22
               Left            =   75
               TabIndex        =   314
               Top             =   270
               Width           =   630
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "COFINS"
            Height          =   1020
            Index           =   1
            Left            =   6570
            TabIndex        =   307
            Top             =   2535
            Width           =   1860
            Begin MSMask.MaskEdBox COFINSRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   308
               Top             =   630
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   21
               Left            =   75
               TabIndex        =   311
               Top             =   705
               Width           =   630
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
               Height          =   195
               Index           =   15
               Left            =   195
               TabIndex        =   310
               Top             =   255
               Width           =   510
            End
            Begin VB.Label COFINSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   750
               TabIndex        =   309
               Top             =   195
               Width           =   915
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "IR"
            Height          =   1356
            Index           =   0
            Left            =   2715
            TabIndex        =   300
            Top             =   2550
            Width           =   1812
            Begin MSMask.MaskEdBox IRAliquota 
               Height          =   285
               Left            =   600
               TabIndex        =   301
               Top             =   600
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorIRRF 
               Height          =   285
               Left            =   600
               TabIndex        =   302
               Top             =   975
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label IRBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   600
               TabIndex        =   306
               Top             =   240
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   29
               Left            =   75
               TabIndex        =   305
               Top             =   285
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Index           =   34
               Left            =   276
               TabIndex        =   304
               Top             =   684
               Width           =   216
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
               Height          =   195
               Index           =   35
               Left            =   75
               TabIndex        =   303
               Top             =   1035
               Width           =   510
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "INSS"
            Height          =   1485
            Index           =   17
            Left            =   150
            TabIndex        =   292
            Top             =   2550
            Width           =   2490
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
               Height          =   285
               Left            =   1155
               TabIndex        =   293
               Top             =   1170
               Width           =   930
            End
            Begin MSMask.MaskEdBox INSSValor 
               Height          =   285
               Left            =   1140
               TabIndex        =   294
               Top             =   885
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox INSSBase 
               Height          =   285
               Left            =   1140
               TabIndex        =   295
               Top             =   210
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox INSSDeducoes 
               Height          =   285
               Left            =   1140
               TabIndex        =   296
               Top             =   555
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
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
               Height          =   195
               Index           =   8
               Left            =   570
               TabIndex        =   299
               Top             =   945
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Left            =   570
               TabIndex        =   298
               Top             =   255
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dedu��es:"
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
               Left            =   150
               TabIndex        =   297
               Top             =   600
               Width           =   930
            End
         End
         Begin VB.CommandButton TributacaoRecalcular 
            Caption         =   "Recalcular Tributa��o"
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
            Left            =   6585
            Style           =   1  'Graphical
            TabIndex        =   291
            Top             =   3570
            Width           =   1815
         End
         Begin VB.Frame Frame10 
            Caption         =   "PIS"
            Height          =   1005
            Index           =   19
            Left            =   4605
            TabIndex        =   285
            Top             =   2520
            Width           =   1860
            Begin MSMask.MaskEdBox PISRetido 
               Height          =   285
               Left            =   780
               TabIndex        =   287
               Top             =   600
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Left            =   105
               TabIndex        =   290
               Top             =   675
               Width           =   630
            End
            Begin VB.Label PISValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   750
               TabIndex        =   289
               Top             =   195
               Width           =   915
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
               Height          =   195
               Index           =   26
               Left            =   195
               TabIndex        =   288
               Top             =   225
               Width           =   510
            End
         End
         Begin MSMask.MaskEdBox TipoTributacao 
            Height          =   330
            Left            =   2145
            TabIndex        =   286
            Top             =   480
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   582
            _Version        =   393216
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
            Left            =   2160
            TabIndex        =   346
            Top             =   45
            Width           =   525
         End
         Begin VB.Label DescTipoTrib 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2775
            TabIndex        =   345
            Top             =   480
            Width           =   5610
         End
         Begin VB.Label LblTipoTrib 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Tributa��o:"
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
            Left            =   420
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   344
            Top             =   555
            Width           =   1695
         End
         Begin VB.Label DescNatOpInterna 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2775
            TabIndex        =   343
            Top             =   30
            Width           =   5610
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
            Left            =   465
            TabIndex        =   342
            Top             =   105
            Width           =   1575
         End
      End
      Begin VB.Frame FrameTributacao 
         BorderStyle     =   0  'None
         Caption         =   "Detalhamento"
         Height          =   4125
         Index           =   2
         Left            =   270
         TabIndex        =   65
         Top             =   390
         Visible         =   0   'False
         Width           =   8670
         Begin VB.Frame Frame2 
            Caption         =   "Sobre"
            Height          =   1305
            Index           =   3
            Left            =   120
            TabIndex        =   121
            Top             =   15
            Width           =   8505
            Begin VB.Frame FrameOutrosTrib 
               Height          =   660
               Left            =   135
               TabIndex        =   124
               Top             =   525
               Visible         =   0   'False
               Width           =   8265
               Begin VB.Label Label20 
                  Caption         =   "Outras Desp.:"
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
                  Index           =   2
                  Left            =   3765
                  TabIndex        =   176
                  Top             =   300
                  Width           =   1185
               End
               Begin VB.Label LabelValorOutrasDespesas 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   4965
                  TabIndex        =   177
                  Top             =   255
                  Width           =   1140
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
                  Height          =   225
                  Index           =   0
                  Left            =   1845
                  TabIndex        =   178
                  Top             =   300
                  Width           =   705
               End
               Begin VB.Label LabelValorSeguro 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   2550
                  TabIndex        =   179
                  Top             =   270
                  Width           =   1140
               End
               Begin VB.Label Label20 
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
                  Height          =   225
                  Index           =   7
                  Left            =   6165
                  TabIndex        =   180
                  Top             =   285
                  Width           =   870
               End
               Begin VB.Label LabelValorDesconto 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   7080
                  TabIndex        =   181
                  Top             =   255
                  Width           =   1140
               End
               Begin VB.Label Label20 
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
                  Height          =   225
                  Index           =   13
                  Left            =   75
                  TabIndex        =   182
                  Top             =   300
                  Width           =   510
               End
               Begin VB.Label LabelValorFrete 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   615
                  TabIndex        =   183
                  Top             =   285
                  Width           =   1140
               End
            End
            Begin VB.Frame FrameItensTrib 
               Caption         =   "Item"
               Height          =   660
               Left            =   120
               TabIndex        =   125
               Top             =   528
               Width           =   8265
               Begin VB.ComboBox ComboItensTrib 
                  Height          =   315
                  Left            =   144
                  Style           =   2  'Dropdown List
                  TabIndex        =   71
                  Top             =   228
                  Width           =   3285
               End
               Begin VB.Label LabelUMItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   7380
                  TabIndex        =   184
                  Top             =   228
                  Width           =   765
               End
               Begin VB.Label LabelQtdeItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   6495
                  TabIndex        =   185
                  Top             =   228
                  Width           =   840
               End
               Begin VB.Label LabelValorItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   4110
                  TabIndex        =   186
                  Top             =   228
                  Width           =   1140
               End
               Begin VB.Label Label19 
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
                  Height          =   225
                  Index           =   1
                  Left            =   3540
                  TabIndex        =   187
                  Top             =   285
                  Width           =   570
               End
               Begin VB.Label Label20 
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
                  Height          =   225
                  Index           =   1
                  Left            =   5385
                  TabIndex        =   188
                  Top             =   285
                  Width           =   1065
               End
            End
            Begin VB.OptionButton TribSobreOutrasDesp 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   4656
               TabIndex        =   69
               Top             =   210
               Width           =   1845
            End
            Begin VB.OptionButton TribSobreSeguro 
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
               Height          =   288
               Left            =   3000
               TabIndex        =   68
               Top             =   210
               Width           =   960
            End
            Begin VB.OptionButton TribSobreDesconto 
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
               Height          =   288
               Left            =   7200
               TabIndex        =   70
               Top             =   225
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.OptionButton TribSobreFrete 
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
               Height          =   288
               Left            =   1488
               TabIndex        =   67
               Top             =   210
               Width           =   816
            End
            Begin VB.OptionButton TribSobreItem 
               Caption         =   "Item"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   108
               TabIndex        =   66
               Top             =   210
               Width           =   684
            End
         End
         Begin VB.Frame Frame15 
            Height          =   2700
            Left            =   120
            TabIndex        =   126
            Top             =   1320
            Width           =   8508
            Begin VB.Frame IPIItemFrame 
               Caption         =   "IPI"
               Height          =   2472
               Left            =   6000
               TabIndex        =   127
               Top             =   180
               Width           =   2376
               Begin VB.CheckBox IPICredita 
                  Caption         =   "Debita"
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
                  Height          =   264
                  Left            =   1008
                  TabIndex        =   88
                  Top             =   2160
                  Width           =   936
               End
               Begin VB.ComboBox ComboIPITipo 
                  Height          =   315
                  Left            =   252
                  Style           =   2  'Dropdown List
                  TabIndex        =   83
                  Top             =   240
                  Width           =   1716
               End
               Begin MSMask.MaskEdBox IPIPercRedBaseItem 
                  Height          =   288
                  Left            =   1272
                  TabIndex        =   85
                  Top             =   1032
                  Width           =   696
                  _ExtentX        =   1217
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIValorItem 
                  Height          =   288
                  Left            =   840
                  TabIndex        =   87
                  Top             =   1836
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.0000"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIAliquotaItem 
                  Height          =   288
                  Left            =   852
                  TabIndex        =   86
                  Top             =   1452
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIBaseItem 
                  Height          =   288
                  Left            =   852
                  TabIndex        =   84
                  Top             =   636
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Red. Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   1
                  Left            =   276
                  TabIndex        =   189
                  Top             =   1104
                  Width           =   888
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Aliq."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   2
                  Left            =   276
                  TabIndex        =   190
                  Top             =   1500
                  Width           =   384
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
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
                  Left            =   240
                  TabIndex        =   191
                  Top             =   1905
                  Width           =   450
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   13
                  Left            =   288
                  TabIndex        =   192
                  Top             =   732
                  Width           =   444
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "ICMS"
               Height          =   1692
               Index           =   2
               Left            =   132
               TabIndex        =   128
               Top             =   960
               Width           =   5688
               Begin VB.CheckBox ICMSCredita 
                  Caption         =   "Debita"
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
                  Height          =   264
                  Left            =   2460
                  TabIndex        =   79
                  Top             =   1368
                  Width           =   936
               End
               Begin VB.Frame Frame2 
                  Caption         =   "Substitui��o"
                  Height          =   1368
                  Index           =   1
                  Left            =   3552
                  TabIndex        =   129
                  Top             =   144
                  Width           =   2004
                  Begin MSMask.MaskEdBox ICMSSubstValorItem 
                     Height          =   288
                     Left            =   672
                     TabIndex        =   82
                     Top             =   984
                     Width           =   1116
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#,##0.0000"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox ICMSSubstAliquotaItem 
                     Height          =   285
                     Left            =   690
                     TabIndex        =   81
                     Top             =   630
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#0.#0\%"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox ICMSSubstBaseItem 
                     Height          =   288
                     Left            =   684
                     TabIndex        =   80
                     Top             =   252
                     Width           =   1092
                     _ExtentX        =   1905
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#,##0.00"
                     PromptChar      =   " "
                  End
                  Begin VB.Label Label19 
                     AutoSize        =   -1  'True
                     Caption         =   "Valor"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   192
                     Index           =   2
                     Left            =   108
                     TabIndex        =   193
                     Top             =   1020
                     Width           =   456
                  End
                  Begin VB.Label Label18 
                     AutoSize        =   -1  'True
                     Caption         =   "Aliq."
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
                     TabIndex        =   194
                     Top             =   675
                     Width           =   390
                  End
                  Begin VB.Label Label20 
                     AutoSize        =   -1  'True
                     Caption         =   "Base"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   192
                     Index           =   21
                     Left            =   120
                     TabIndex        =   195
                     Top             =   312
                     Width           =   444
                  End
               End
               Begin VB.ComboBox ComboICMSTipo 
                  Height          =   315
                  Left            =   135
                  Style           =   2  'Dropdown List
                  TabIndex        =   74
                  Top             =   228
                  Width           =   3336
               End
               Begin MSMask.MaskEdBox ICMSValorItem 
                  Height          =   288
                  Left            =   2316
                  TabIndex        =   78
                  Top             =   1008
                  Width           =   1116
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.0000"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSAliquotaItem 
                  Height          =   288
                  Left            =   2316
                  TabIndex        =   77
                  Top             =   624
                  Width           =   1116
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSPercRedBaseItem 
                  Height          =   288
                  Left            =   1032
                  TabIndex        =   76
                  Top             =   1008
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSBaseItem 
                  Height          =   288
                  Left            =   588
                  TabIndex        =   75
                  Top             =   624
                  Width           =   1116
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Red. Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   1
                  Left            =   96
                  TabIndex        =   196
                  Top             =   1068
                  Width           =   888
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Aliq."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   2
                  Left            =   1812
                  TabIndex        =   197
                  Top             =   648
                  Width           =   384
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
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
                  Left            =   1800
                  TabIndex        =   198
                  Top             =   1050
                  Width           =   450
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   1
                  Left            =   84
                  TabIndex        =   199
                  Top             =   648
                  Width           =   444
               End
            End
            Begin MSMask.MaskEdBox NaturezaOpItem 
               Height          =   300
               Left            =   1824
               TabIndex        =   72
               Top             =   204
               Width           =   552
               _ExtentX        =   953
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
            Begin MSMask.MaskEdBox TipoTributacaoItem 
               Height          =   300
               Left            =   1824
               TabIndex        =   73
               Top             =   612
               Width           =   552
               _ExtentX        =   953
               _ExtentY        =   529
               _Version        =   393216
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
            Begin VB.Label DescTipoTribItem 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2460
               TabIndex        =   200
               Top             =   648
               Width           =   3165
            End
            Begin VB.Label LabelDescrNatOpItem 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2460
               TabIndex        =   201
               Top             =   210
               Width           =   3165
            End
            Begin VB.Label NaturezaItemLabel 
               AutoSize        =   -1  'True
               Caption         =   "Natureza Oper.:"
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
               Left            =   420
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   202
               Top             =   255
               Width           =   1365
            End
            Begin VB.Label LblTipoTribItem 
               Caption         =   "Tipo de Tributa��o:"
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
               Height          =   225
               Left            =   90
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   203
               Top             =   675
               Width           =   1710
            End
         End
      End
      Begin MSComctlLib.TabStrip OpcaoTributacao 
         Height          =   4590
         Left            =   180
         TabIndex        =   130
         Top             =   15
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   8096
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Resumo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhamento"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4725
      Index           =   5
      Left            =   150
      TabIndex        =   44
      Top             =   1425
      Visible         =   0   'False
      Width           =   9195
      Begin VB.CheckBox ComissaoAutomatica 
         Caption         =   "Calcula comiss�o automaticamente"
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
         Left            =   255
         TabIndex        =   45
         Top             =   120
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comiss�es"
         Height          =   4200
         Index           =   0
         Left            =   75
         TabIndex        =   114
         Top             =   420
         Width           =   9075
         Begin VB.Frame SSFrame4 
            Caption         =   "Totais - Comiss�es"
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
            TabIndex        =   277
            Top             =   3120
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
               TabIndex        =   283
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
               TabIndex        =   282
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label TotalPercentualComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3840
               TabIndex        =   281
               Top             =   360
               Width           =   855
            End
            Begin VB.Label TotalValorComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5520
               TabIndex        =   280
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
               TabIndex        =   279
               Top             =   360
               Width           =   990
            End
            Begin VB.Label TotalValorBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1200
               TabIndex        =   278
               Top             =   360
               Width           =   1215
            End
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
            Height          =   555
            Left            =   7380
            Picture         =   "nfiscalgr.ctx":0536
            Style           =   1  'Graphical
            TabIndex        =   276
            Top             =   3285
            Width           =   1650
         End
         Begin VB.ComboBox DiretoIndireto 
            Height          =   315
            ItemData        =   "nfiscalgr.ctx":0AE0
            Left            =   6405
            List            =   "nfiscalgr.ctx":0AEA
            Style           =   2  'Dropdown List
            TabIndex        =   275
            Top             =   1065
            Width           =   1335
         End
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   225
            Left            =   3645
            TabIndex        =   49
            Top             =   255
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   225
            Left            =   2550
            TabIndex        =   48
            Top             =   255
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   225
            Left            =   1695
            TabIndex        =   47
            Top             =   255
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
            Left            =   435
            TabIndex        =   46
            Top             =   255
            Width           =   1260
            _ExtentX        =   2223
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
            Left            =   5610
            TabIndex        =   51
            Top             =   255
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
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   225
            Left            =   4740
            TabIndex        =   50
            Top             =   255
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox ValorBaixa 
            Height          =   225
            Left            =   7470
            TabIndex        =   53
            Top             =   255
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
            Left            =   6705
            TabIndex        =   52
            Top             =   255
            Width           =   750
            _ExtentX        =   1323
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
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   2655
            Left            =   90
            TabIndex        =   54
            Top             =   330
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   4683
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
      Caption         =   "Contabilizacao"
      Height          =   4725
      Index           =   8
      Left            =   150
      TabIndex        =   89
      Top             =   1425
      Visible         =   0   'False
      Width           =   9195
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
         Index           =   0
         Left            =   6360
         TabIndex        =   95
         Top             =   375
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
         Index           =   0
         Left            =   6360
         TabIndex        =   93
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   900
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
         Left            =   7830
         TabIndex        =   94
         Top             =   60
         Width           =   1245
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descri��o do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   131
         Top             =   3495
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   204
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   205
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
            TabIndex        =   206
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
            TabIndex        =   207
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2790
         Left            =   6375
         TabIndex        =   107
         Top             =   1590
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   104
         Top             =   1635
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   105
         Top             =   2025
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
         Left            =   3465
         TabIndex        =   98
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4665
         TabIndex        =   103
         Top             =   1335
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
         TabIndex        =   99
         Top             =   1320
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
         TabIndex        =   102
         Top             =   1350
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
         TabIndex        =   101
         Top             =   1290
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
         TabIndex        =   100
         Top             =   1335
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
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   570
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   585
         TabIndex        =   92
         Top             =   540
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
         Left            =   5610
         TabIndex        =   91
         Top             =   150
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
         TabIndex        =   90
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
         TabIndex        =   106
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
         Left            =   6360
         TabIndex        =   108
         Top             =   1575
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
         Left            =   6345
         TabIndex        =   109
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
         Index           =   0
         Left            =   6390
         TabIndex        =   96
         Top             =   720
         Width           =   690
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
         Left            =   6360
         TabIndex        =   208
         Top             =   1290
         Visible         =   0   'False
         Width           =   2490
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
         TabIndex        =   209
         Top             =   180
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
         TabIndex        =   210
         Top             =   180
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
         TabIndex        =   211
         Top             =   570
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   212
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   213
         Top             =   3045
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
         TabIndex        =   214
         Top             =   3060
         Width           =   615
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
         TabIndex        =   215
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
         TabIndex        =   216
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
         TabIndex        =   217
         Top             =   960
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
         TabIndex        =   218
         Top             =   600
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   219
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   220
         Top             =   585
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
         TabIndex        =   221
         Top             =   615
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   222
         Top             =   135
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
         Height          =   255
         Left            =   45
         TabIndex        =   223
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Rastro"
      Height          =   4605
      Index           =   9
      Left            =   150
      TabIndex        =   224
      Top             =   1425
      Visible         =   0   'False
      Width           =   9195
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
         TabIndex        =   227
         Top             =   4170
         Width           =   1665
      End
      Begin VB.Frame Frame18 
         Caption         =   "Rastreamento do Produto"
         Height          =   4050
         Left            =   30
         TabIndex        =   225
         Top             =   120
         Width           =   9030
         Begin VB.ComboBox EscaninhoRastro 
            Height          =   315
            ItemData        =   "nfiscalgr.ctx":0B00
            Left            =   3690
            List            =   "nfiscalgr.ctx":0B0D
            Style           =   2  'Dropdown List
            TabIndex        =   237
            Top             =   285
            Width           =   1215
         End
         Begin MSMask.MaskEdBox UMRastro 
            Height          =   240
            Left            =   3075
            TabIndex        =   228
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
            TabIndex        =   229
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
            TabIndex        =   230
            Top             =   285
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
            TabIndex        =   231
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
         Begin MSMask.MaskEdBox ProdutoRastro 
            Height          =   240
            Left            =   360
            TabIndex        =   232
            Top             =   450
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteRastro 
            Height          =   225
            Left            =   2820
            TabIndex        =   233
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteDataRastro 
            Height          =   255
            Left            =   5580
            TabIndex        =   234
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
            TabIndex        =   235
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
            TabIndex        =   236
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
            TabIndex        =   226
            Top             =   225
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
      Caption         =   "Frame5"
      Height          =   4725
      Index           =   6
      Left            =   150
      TabIndex        =   55
      Top             =   1425
      Visible         =   0   'False
      Width           =   9195
      Begin VB.CheckBox ImprimeRomaneio 
         Caption         =   "Imprime Romaneio na Grava��o."
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
         Left            =   390
         TabIndex        =   112
         Top             =   195
         Width           =   3270
      End
      Begin VB.CommandButton BotaoLocalizacao 
         Caption         =   "Localiza��o de Produto"
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
         Left            =   5895
         TabIndex        =   63
         Top             =   4230
         Width           =   2715
      End
      Begin VB.Frame Frame10 
         Caption         =   "Localiza��o dos Produtos"
         Height          =   3465
         Index           =   14
         Left            =   255
         TabIndex        =   113
         Top             =   615
         Width           =   8355
         Begin MSMask.MaskEdBox UnidadeMedEst 
            Height          =   225
            Left            =   7410
            TabIndex        =   61
            Top             =   690
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
            Left            =   1635
            TabIndex        =   57
            Top             =   480
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
            Left            =   2985
            TabIndex        =   58
            Top             =   480
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
            Left            =   4320
            TabIndex        =   59
            Top             =   480
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
            Left            =   1080
            TabIndex        =   56
            Top             =   480
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
            Left            =   5925
            TabIndex        =   60
            Top             =   675
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
            Height          =   2910
            Left            =   360
            TabIndex        =   62
            Top             =   360
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   5133
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
      Left            =   6555
      ScaleHeight     =   465
      ScaleWidth      =   2745
      TabIndex        =   238
      TabStop         =   0   'False
      Top             =   60
      Width           =   2805
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   1357
         Picture         =   "nfiscalgr.ctx":0B2C
         Style           =   1  'Graphical
         TabIndex        =   243
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   390
      End
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
         Left            =   60
         Picture         =   "nfiscalgr.ctx":0CB6
         Style           =   1  'Graphical
         TabIndex        =   242
         ToolTipText     =   "Consulta de T�tulo a Receber"
         Top             =   75
         Width           =   765
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   896
         Picture         =   "nfiscalgr.ctx":1538
         Style           =   1  'Graphical
         TabIndex        =   241
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   1818
         Picture         =   "nfiscalgr.ctx":1692
         Style           =   1  'Graphical
         TabIndex        =   240
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   2280
         Picture         =   "nfiscalgr.ctx":1BC4
         Style           =   1  'Graphical
         TabIndex        =   239
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   390
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5460
      Left            =   90
      TabIndex        =   110
      Top             =   720
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   9631
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comprovante"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comiss�es"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Almoxarifado"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributa��o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabiliza��o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "NFiscalGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTNFiscal
Attribute objCT.VB_VarHelpID = -1

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

Private Sub BotaoComprovante_Click()
    Call objCT.BotaoComprovante_Click
End Sub

Private Sub ComprovServCon_Change()
    Call objCT.ComprovServCon_Change
End Sub

Private Sub ComprovServCon_GotFocus()
    Call objCT.ComprovServCon_GotFocus
End Sub

Private Sub ComprovServCon_KeyPress(KeyAscii As Integer)
    Call objCT.ComprovServCon_KeyPress(KeyAscii)
End Sub

Private Sub ComprovServCon_Validate(Cancel As Boolean)
    Call objCT.ComprovServCon_Validate(Cancel)
End Sub

Private Sub GridComprovServ_Click()
    Call objCT.GridComprovServ_Click
End Sub

Private Sub GridComprovServ_EnterCell()
    Call objCT.GridComprovServ_EnterCell
End Sub

Private Sub GridComprovServ_GotFocus()
    Call objCT.GridComprovServ_GotFocus
End Sub

Private Sub GridComprovServ_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.GridComprovServ_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridComprovServ_KeyPress(KeyAscii As Integer)
    Call objCT.GridComprovServ_KeyPress(KeyAscii)
End Sub

Private Sub GridComprovServ_LeaveCell()
    Call objCT.GridComprovServ_LeaveCell
End Sub

Private Sub GridComprovServ_Scroll()
    Call objCT.GridComprovServ_Scroll
End Sub

Private Sub GridComprovServ_RowColChange()
    Call objCT.GridComprovServ_RowColChange
End Sub

Private Sub GridComprovServ_Validate(Cancel As Boolean)
    Call objCT.GridComprovServ_Validate(Cancel)
End Sub

Private Sub ISSValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ISSValor1, Source, X, Y)
End Sub

Private Sub ISSValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ISSValor1, Button, Shift, X, Y)
End Sub

Private Sub ValorContainer_Change()
    Call objCT.ValorContainer_Change
End Sub

Private Sub ValorMercadoria_Change()
    Call objCT.ValorMercadoria_Change
End Sub

Private Sub Destino_Change()
    Call objCT.Destino_Change
End Sub

Private Sub Origem_Change()
    Call objCT.Origem_Change
End Sub

Private Sub BotaoConsultaTitRec_Click()
    Call objCT.BotaoConsultaTitRec_Click
End Sub

Private Sub BotaoExcluir_Click()
    Call objCT.BotaoExcluir_Click
End Sub

'Private Sub BotaoRastreamento_Click()
'    Call objCT.BotaoRastreamento_Click
'End Sub
'
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

Private Sub AlmoxRastro_Change()
     Call objCT.AlmoxRastro_Change
End Sub

Private Sub AlmoxRastro_GotFocus()
     Call objCT.AlmoxRastro_GotFocus
End Sub

Private Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
     Call objCT.AlmoxRastro_KeyPress(KeyAscii)
End Sub

Private Sub AlmoxRastro_Validate(Cancel As Boolean)
     Call objCT.AlmoxRastro_Validate(Cancel)
End Sub

Private Sub EscaninhoRastro_Change()
     Call objCT.QuantLoteRastro_Change
End Sub

Private Sub EscaninhoRastro_GotFocus()
     Call objCT.EscaninhoRastro_GotFocus
End Sub

Private Sub EscaninhoRastro_KeyPress(KeyAscii As Integer)
     Call objCT.EscaninhoRastro_KeyPress(KeyAscii)
End Sub

Private Sub EscaninhoRastro_Validate(Cancel As Boolean)
     Call objCT.EscaninhoRastro_Validate(Cancel)
End Sub

'Fim Rastreamento

Private Sub UserControl_Initialize()
    Set objCT = New CTNFiscal
    Set objCT.objUserControl = Me
End Sub

Private Sub ImprimeRomaneio_Click()
    Call objCT.ImprimeRomaneio_Click
End Sub

Private Sub LblTipoTrib_Click()
    Call objCT.LblTipoTrib_Click
End Sub

Private Sub LblTipoTribItem_Click()
    Call objCT.LblTipoTribItem_Click
End Sub

Private Sub BotaoLimparNF_Click()
     Call objCT.BotaoLimparNF_Click
End Sub

Private Sub ComboICMSTipo_Click()
     Call objCT.ComboICMSTipo_Click
End Sub

Private Sub ComboIPITipo_Click()
     Call objCT.ComboIPITipo_Click
End Sub

Private Sub ComboItensTrib_Click()
     Call objCT.ComboItensTrib_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub ComissaoAutomatica_Click()
     Call objCT.ComissaoAutomatica_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataSaida_GotFocus()
     Call objCT.DataSaida_GotFocus
End Sub

Private Sub Destinatario_Click()
     Call objCT.Destinatario_Click
End Sub

Private Sub Emitente_Click()
     Call objCT.Emitente_Click
End Sub

Private Sub LblNatOpInterna_Click()
     Call objCT.LblNatOpInterna_Click
End Sub

Private Sub NaturezaItemLabel_Click()
     Call objCT.NaturezaItemLabel_Click
End Sub

Private Sub NaturezaOpItem_Change()
     Call objCT.NaturezaOpItem_Change
End Sub

Private Sub NaturezaOpItem_Validate(Cancel As Boolean)
     Call objCT.NaturezaOpItem_Validate(Cancel)
End Sub

Private Sub NatOpInterna_Change()
     Call objCT.NatOpInterna_Change
End Sub

Private Sub NatOpInterna_Validate(Cancel As Boolean)
     Call objCT.NatOpInterna_Validate(Cancel)
End Sub

Private Sub NFiscal_GotFocus()
     Call objCT.NFiscal_GotFocus
End Sub

Private Sub NFiscalOriginal_GotFocus()
     Call objCT.NFiscalOriginal_GotFocus
End Sub

Private Sub SerieNFOriginalLabel_Click()
     Call objCT.SerieNFOriginalLabel_Click
End Sub

Private Sub TipoTributacao_Change()
     Call objCT.TipoTributacao_Change
End Sub

Private Sub TipoTributacao_Validate(Cancel As Boolean)
     Call objCT.TipoTributacao_Validate(Cancel)
End Sub

Private Sub TipoTributacaoItem_Change()
     Call objCT.TipoTributacaoItem_Change
End Sub

Private Sub TipoTributacaoItem_Validate(Cancel As Boolean)
     Call objCT.TipoTributacaoItem_Validate(Cancel)
End Sub

Private Sub NaturezaOpItem_GotFocus()
     Call objCT.NaturezaOpItem_GotFocus
End Sub

Private Sub NatOpInterna_GotFocus()
     Call objCT.NatOpInterna_GotFocus
End Sub

Private Sub TipoTributacao_GotFocus()
     Call objCT.TipoTributacao_GotFocus
End Sub

Private Sub TipoTributacaoItem_GotFocus()
     Call objCT.TipoTributacaoItem_GotFocus
End Sub

Private Sub TribSobreDesconto_Click()
     Call objCT.TribSobreDesconto_Click
End Sub

Private Sub TribSobreFrete_Click()
     Call objCT.TribSobreFrete_Click
End Sub

Private Sub TribSobreItem_Click()
     Call objCT.TribSobreItem_Click
End Sub

Private Sub TribSobreOutrasDesp_Click()
     Call objCT.TribSobreOutrasDesp_Click
End Sub

Private Sub TribSobreSeguro_Click()
     Call objCT.TribSobreSeguro_Click
End Sub

Private Sub TributacaoRecalcular_Click()
     Call objCT.TributacaoRecalcular_Click
End Sub

Private Sub OpcaoTributacao_Click()
     Call objCT.OpcaoTributacao_Click
End Sub

Private Sub ValorIRRF_Change()
     Call objCT.ValorIRRF_Change
End Sub

Private Sub ICMSAliquotaItem_Change()
     Call objCT.ICMSAliquotaItem_Change
End Sub

Private Sub ICMSAliquotaItem_Validate(Cancel As Boolean)
     Call objCT.ICMSAliquotaItem_Validate(Cancel)
End Sub

Private Sub ICMSBaseItem_Change()
     Call objCT.ICMSBaseItem_Change
End Sub

Private Sub ICMSBaseItem_Validate(Cancel As Boolean)
     Call objCT.ICMSBaseItem_Validate(Cancel)
End Sub

Private Sub ICMSPercRedBaseItem_Change()
     Call objCT.ICMSPercRedBaseItem_Change
End Sub

Private Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)
     Call objCT.ICMSPercRedBaseItem_Validate(Cancel)
End Sub

Private Sub ICMSSubstAliquotaItem_Change()
     Call objCT.ICMSSubstAliquotaItem_Change
End Sub

Private Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)
     Call objCT.ICMSSubstAliquotaItem_Validate(Cancel)
End Sub

Private Sub ICMSSubstBaseItem_Change()
     Call objCT.ICMSSubstBaseItem_Change
End Sub

Private Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)
     Call objCT.ICMSSubstBaseItem_Validate(Cancel)
End Sub

Private Sub ICMSSubstValorItem_Change()
     Call objCT.ICMSSubstValorItem_Change
End Sub

Private Sub ICMSSubstValorItem_Validate(Cancel As Boolean)
     Call objCT.ICMSSubstValorItem_Validate(Cancel)
End Sub

Private Sub ICMSValorItem_Change()
     Call objCT.ICMSValorItem_Change
End Sub

Private Sub ICMSValorItem_Validate(Cancel As Boolean)
     Call objCT.ICMSValorItem_Validate(Cancel)
End Sub

Private Sub IPIAliquotaItem_Change()
     Call objCT.IPIAliquotaItem_Change
End Sub

Private Sub IPIAliquotaItem_Validate(Cancel As Boolean)
     Call objCT.IPIAliquotaItem_Validate(Cancel)
End Sub

Private Sub IPIBaseItem_Change()
     Call objCT.IPIBaseItem_Change
End Sub

Private Sub IPIBaseItem_Validate(Cancel As Boolean)
     Call objCT.IPIBaseItem_Validate(Cancel)
End Sub

Private Sub IPIPercRedBaseItem_Change()
     Call objCT.IPIPercRedBaseItem_Change
End Sub

Private Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)
     Call objCT.IPIPercRedBaseItem_Validate(Cancel)
End Sub

Private Sub IPIValorItem_Change()
     Call objCT.IPIValorItem_Change
End Sub

Private Sub IPIValorItem_Validate(Cancel As Boolean)
     Call objCT.IPIValorItem_Validate(Cancel)
End Sub

Private Sub IRAliquota_Change()
     Call objCT.IRAliquota_Change
End Sub

Private Sub IRAliquota_Validate(Cancel As Boolean)
     Call objCT.IRAliquota_Validate(Cancel)
End Sub

Private Sub ISSAliquota_Change()
     Call objCT.ISSAliquota_Change
End Sub

Private Sub ISSAliquota_Validate(Cancel As Boolean)
     Call objCT.ISSAliquota_Validate(Cancel)
End Sub

Private Sub ISSIncluso_Click()
     Call objCT.ISSIncluso_Click
End Sub

Private Sub ISSValor_Change()
     Call objCT.ISSValor_Change
End Sub

Private Sub ISSValor_Validate(Cancel As Boolean)
     Call objCT.ISSValor_Validate(Cancel)
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

Private Sub Almox_Change()
     Call objCT.Almox_Change
End Sub

Private Sub CanalVenda_Change()
     Call objCT.CanalVenda_Change
End Sub

Private Sub CanalVenda_Click()
     Call objCT.CanalVenda_Click
End Sub

Private Sub CanalVenda_Validate(Cancel As Boolean)
     Call objCT.CanalVenda_Validate(Cancel)
End Sub

Private Sub ClienteLabel_Click()
     Call objCT.ClienteLabel_Click
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
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

Private Sub DataSaida_Change()
     Call objCT.DataSaida_Change
End Sub

Private Sub DataSaida_Validate(Cancel As Boolean)
     Call objCT.DataSaida_Validate(Cancel)
End Sub

'horasaida
Private Sub HoraSaida_Change()
     Call objCT.HoraSaida_Change
End Sub

'horasaida
Private Sub HoraSaida_Validate(Cancel As Boolean)
     Call objCT.HoraSaida_Validate(Cancel)
End Sub

'horasaida
Private Sub HoraSaida_GotFocus()
     Call objCT.HoraSaida_GotFocus
End Sub

Private Sub UpDownSaida_DownClick()
     Call objCT.UpDownSaida_DownClick
End Sub

Private Sub UpDownSaida_UpClick()
     Call objCT.UpDownSaida_UpClick
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

Private Sub DescricaoItem_Change()
     Call objCT.DescricaoItem_Change
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

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub FilialEntrega_Change()
     Call objCT.FilialEntrega_Change
End Sub

Private Sub FilialEntrega_Click()
     Call objCT.FilialEntrega_Click
End Sub

Private Sub FilialEntrega_Validate(Cancel As Boolean)
     Call objCT.FilialEntrega_Validate(Cancel)
End Sub

Private Sub ItemNFiscal_Change()
     Call objCT.ItemNFiscal_Change
End Sub

Private Sub Mensagem_Change()
     Call objCT.Mensagem_Change
End Sub

Private Sub NFiscal_Change()
     Call objCT.NFiscal_Change
End Sub

Private Sub NFiscalLabel_Click()
     Call objCT.NFiscalLabel_Click
End Sub

Private Sub NFiscalOriginal_Change()
     Call objCT.NFiscalOriginal_Change
End Sub

Private Sub NumPedidoTerc_Change()
     Call objCT.NumPedidoTerc_Change
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

Private Sub PercentualBaixa_Change()
     Call objCT.PercentualBaixa_Change
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

Private Sub PesoBruto_Change()
     Call objCT.PesoBruto_Change
End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)
     Call objCT.PesoBruto_Validate(Cancel)
End Sub

Private Sub PesoLiquido_Change()
     Call objCT.PesoLiquido_Change
End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)
     Call objCT.PesoLiquido_Validate(Cancel)
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

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub ProdutoAlmox_Change()
     Call objCT.ProdutoAlmox_Change
End Sub

Private Sub QuantAlocada_Change()
     Call objCT.QuantAlocada_Change
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

Private Sub QuantVendida_Change()
     Call objCT.QuantVendida_Change
End Sub

Private Sub Serie_Change()
     Call objCT.Serie_Change
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

Private Sub SerieNFiscalOriginal_Change()
     Call objCT.SerieNFiscalOriginal_Change
End Sub

Private Sub SerieNFiscalOriginal_Click()
     Call objCT.SerieNFiscalOriginal_Click
End Sub

Private Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)
     Call objCT.SerieNFiscalOriginal_Validate(Cancel)
End Sub

Private Sub TabelaPreco_Change()
     Call objCT.TabelaPreco_Change
End Sub

Private Sub TabelaPreco_Click()
     Call objCT.TabelaPreco_Click
End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)
     Call objCT.TabelaPreco_Validate(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub TipoNFiscal_Change()
     Call objCT.TipoNFiscal_Change
End Sub

Private Sub TipoNFiscal_Click()
     Call objCT.TipoNFiscal_Click
End Sub

Private Sub TipoNFiscal_Validate(Cancel As Boolean)
     Call objCT.TipoNFiscal_Validate(Cancel)
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub Transportadora_Click()
     Call objCT.Transportadora_Click
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub UnidadeMed_Change()
     Call objCT.UnidadeMed_Change
End Sub

Private Sub UnidadeMed_Click()
     Call objCT.UnidadeMed_Click
End Sub

Private Sub UnidadeMedEst_Change()
     Call objCT.UnidadeMedEst_Change
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

Private Sub ValorBaixa_Change()
     Call objCT.ValorBaixa_Change
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

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub ValorDespesas_Change()
     Call objCT.ValorDespesas_Change
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

Private Sub ValorFrete_Change()
     Call objCT.ValorFrete_Change
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

Private Sub VolumeEspecie_Change()
     Call objCT.VolumeEspecie_Change
End Sub

Private Sub VolumeMarca_Change()
     Call objCT.VolumeMarca_Change
End Sub

Private Sub VolumeNumero_Change()
     Call objCT.VolumeNumero_Change
End Sub

Private Sub VolumeQuant_Change()
     Call objCT.VolumeQuant_Change
End Sub

Private Sub BotaoEstoqueProd_Click()
     Call objCT.BotaoEstoqueProd_Click
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub NFiscalOriginalLabel_Click()
     Call objCT.NFiscalOriginalLabel_Click
End Sub

Private Sub BotaoVendedores_Click()
     Call objCT.BotaoVendedores_Click
End Sub

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long
     Trata_Parametros = objCT.Trata_Parametros(objNFiscal)
End Function

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

Private Sub ValorFrete_Validate(Cancel As Boolean)
     Call objCT.ValorFrete_Validate(Cancel)
End Sub

Private Sub ValorIRRF_Validate(Cancel As Boolean)
     Call objCT.ValorIRRF_Validate(Cancel)
End Sub

Private Sub GridComissoes_Click()
     Call objCT.GridComissoes_Click
End Sub

Private Sub GridComissoes_EnterCell()
     Call objCT.GridComissoes_EnterCell
End Sub

Private Sub GridComissoes_GotFocus()
     Call objCT.GridComissoes_GotFocus
End Sub

Private Sub GridComissoes_KeyPress(KeyAscii As Integer)
     Call objCT.GridComissoes_KeyPress(KeyAscii)
End Sub

Private Sub GridComissoes_LeaveCell()
     Call objCT.GridComissoes_LeaveCell
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

Private Sub ValorDespesas_Validate(Cancel As Boolean)
     Call objCT.ValorDespesas_Validate(Cancel)
End Sub

Private Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridComissoes_KeyDown(KeyCode, Shift)
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

Private Sub BotaoLocalizacao_Click()
     Call objCT.BotaoLocalizacao_Click
End Sub

''''Private Sub CTBBotaoModeloPadrao_Click()
''''     Call objCT.CTBBotaoModeloPadrao_Click
''''End Sub

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
'''
'''Private Sub CTBBotaoLimparGrid_Click()
'''     Call objCT.CTBBotaoLimparGrid_Click
'''End Sub

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

Private Sub Filial_Formata(objFilial As Object, iFilial As Integer)
     Call objCT.Filial_Formata(objFilial, iFilial)
End Sub

Private Sub MensagemLabel_Click()
     Call objCT.MensagemLabel_Click
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        Set objCT.objUserControl = Nothing
        Set objCT = Nothing
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

Private Sub Label7_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label7(Index), Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7(Index), Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label17(Index), Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17(Index), Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label15(Index), Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label18(Index), Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18(Index), Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label13(Index), Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13(Index), Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label16(Index), Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16(Index), Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label30(Index), Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30(Index), Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label19(Index), Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19(Index), Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label6(Index), Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6(Index), Button, Shift, X, Y)
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

Private Sub LblNatOpInterna_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNatOpInterna, Source, X, Y)
End Sub

Private Sub LblNatOpInterna_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNatOpInterna, Button, Shift, X, Y)
End Sub

Private Sub SerieLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieLabel, Source, X, Y)
End Sub

Private Sub SerieLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieLabel, Button, Shift, X, Y)
End Sub

Private Sub NFiscalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalLabel, Source, X, Y)
End Sub

Private Sub NFiscalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalLabel, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
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
'
'Private Sub LabelTotaisComissoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelTotaisComissoes, Button, Shift, X, Y)
'End Sub

Private Sub ICMSSubstValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstValor1, Source, X, Y)
End Sub

Private Sub ICMSSubstValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstValor1, Button, Shift, X, Y)
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

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
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

Private Sub NFiscalOriginalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalOriginalLabel, Source, X, Y)
End Sub

Private Sub NFiscalOriginalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalOriginalLabel, Button, Shift, X, Y)
End Sub

Private Sub SerieNFOriginalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieNFOriginalLabel, Source, X, Y)
End Sub

Private Sub SerieNFOriginalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieNFOriginalLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelValorOutrasDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorOutrasDespesas, Source, X, Y)
End Sub

Private Sub LabelValorOutrasDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorOutrasDespesas, Button, Shift, X, Y)
End Sub

Private Sub LabelValorSeguro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorSeguro, Source, X, Y)
End Sub

Private Sub LabelValorSeguro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorSeguro, Button, Shift, X, Y)
End Sub

Private Sub LabelValorDesconto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorDesconto, Source, X, Y)
End Sub

Private Sub LabelValorDesconto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorDesconto, Button, Shift, X, Y)
End Sub

Private Sub LabelValorFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorFrete, Source, X, Y)
End Sub

Private Sub LabelValorFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorFrete, Button, Shift, X, Y)
End Sub

Private Sub LabelUMItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelUMItem, Source, X, Y)
End Sub

Private Sub LabelUMItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelUMItem, Button, Shift, X, Y)
End Sub

Private Sub LabelQtdeItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelQtdeItem, Source, X, Y)
End Sub

Private Sub LabelQtdeItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelQtdeItem, Button, Shift, X, Y)
End Sub

Private Sub LabelValorItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorItem, Source, X, Y)
End Sub

Private Sub LabelValorItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorItem, Button, Shift, X, Y)
End Sub

Private Sub DescTipoTribItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescTipoTribItem, Source, X, Y)
End Sub

Private Sub DescTipoTribItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescTipoTribItem, Button, Shift, X, Y)
End Sub

Private Sub LabelDescrNatOpItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescrNatOpItem, Source, X, Y)
End Sub

Private Sub LabelDescrNatOpItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescrNatOpItem, Button, Shift, X, Y)
End Sub

Private Sub NaturezaItemLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NaturezaItemLabel, Source, X, Y)
End Sub

Private Sub NaturezaItemLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NaturezaItemLabel, Button, Shift, X, Y)
End Sub

Private Sub LblTipoTribItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipoTribItem, Source, X, Y)
End Sub

Private Sub LblTipoTribItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipoTribItem, Button, Shift, X, Y)
End Sub

Private Sub ISSBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ISSBase, Source, X, Y)
End Sub

Private Sub ISSBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ISSBase, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstValor, Source, X, Y)
End Sub

Private Sub ICMSSubstValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstValor, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstBase, Source, X, Y)
End Sub

Private Sub ICMSSubstBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstBase, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label20(Index), Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20(Index), Button, Shift, X, Y)
End Sub

Private Sub ICMSCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSCredito, Source, X, Y)
End Sub

Private Sub ICMSCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSCredito, Button, Shift, X, Y)
End Sub

Private Sub ICMSValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSValor, Source, X, Y)
End Sub

Private Sub ICMSValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSValor, Button, Shift, X, Y)
End Sub

Private Sub ICMSBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSBase, Source, X, Y)
End Sub

Private Sub ICMSBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSBase, Button, Shift, X, Y)
End Sub

Private Sub IPIValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor, Source, X, Y)
End Sub

Private Sub IPIValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor, Button, Shift, X, Y)
End Sub

Private Sub IPIBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIBase, Source, X, Y)
End Sub

Private Sub IPIBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIBase, Button, Shift, X, Y)
End Sub

Private Sub IPICredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPICredito, Source, X, Y)
End Sub

Private Sub IPICredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPICredito, Button, Shift, X, Y)
End Sub

Private Sub IRBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IRBase, Source, X, Y)
End Sub

Private Sub IRBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IRBase, Button, Shift, X, Y)
End Sub

Private Sub LblNatOpInternaEspelho_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNatOpInternaEspelho, Source, X, Y)
End Sub

Private Sub LblNatOpInternaEspelho_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNatOpInternaEspelho, Button, Shift, X, Y)
End Sub

Private Sub DescNatOpInterna_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescNatOpInterna, Source, X, Y)
End Sub

Private Sub DescNatOpInterna_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescNatOpInterna, Button, Shift, X, Y)
End Sub

Private Sub LblTipoTrib_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipoTrib, Source, X, Y)
End Sub

Private Sub LblTipoTrib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipoTrib, Button, Shift, X, Y)
End Sub

Private Sub DescTipoTrib_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescTipoTrib, Source, X, Y)
End Sub

Private Sub DescTipoTrib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescTipoTrib, Button, Shift, X, Y)
End Sub

Private Sub NatOpInternaEspelho_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NatOpInternaEspelho, Source, X, Y)
End Sub

Private Sub NatOpInternaEspelho_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NatOpInternaEspelho, Button, Shift, X, Y)
End Sub

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

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
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
'''
'''Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
'''   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
'''End Sub
'''
'''Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
'''End Sub

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

Private Sub OpcaoTributacao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, OpcaoTributacao)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

'jones-15/03/01
Private Sub INSSBase_Change()
    Call objCT.INSSBase_Change
End Sub

Private Sub INSSBase_Validate(Cancel As Boolean)
    Call objCT.INSSBase_Validate(Cancel)
End Sub

Private Sub INSSDeducoes_Change()
    Call objCT.INSSDeducoes_Change
End Sub

Private Sub INSSDeducoes_Validate(Cancel As Boolean)
    Call objCT.INSSDeducoes_Validate(Cancel)
End Sub

Private Sub INSSRetido_Click()
    Call objCT.INSSRetido_Click
End Sub

Private Sub INSSValor_Change()
    Call objCT.INSSValor_Change
End Sub

Private Sub INSSValor_Validate(Cancel As Boolean)
    Call objCT.INSSValor_Validate(Cancel)
End Sub
'fim jones-15/03/01

Private Sub TotalValorBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorBase, Source, X, Y)
End Sub

Private Sub TotalValorBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorBase, Button, Shift, X, Y)
End Sub

Private Sub PISRetido_Change()
     Call objCT.PISRetido_Change
End Sub

Private Sub PISRetido_Validate(Cancel As Boolean)
     Call objCT.PISRetido_Validate(Cancel)
End Sub

Private Sub COFINSRetido_Change()
     Call objCT.COFINSRetido_Change
End Sub

Private Sub COFINSRetido_Validate(Cancel As Boolean)
     Call objCT.COFINSRetido_Validate(Cancel)
End Sub

Private Sub CSLLRetido_Change()
     Call objCT.CSLLRetido_Change
End Sub

Private Sub CSLLRetido_Validate(Cancel As Boolean)
     Call objCT.CSLLRetido_Validate(Cancel)
End Sub




VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.UserControl LojaConfig 
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   ScaleHeight     =   4905
   ScaleWidth      =   9420
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   3990
      Index           =   3
      Left            =   135
      TabIndex        =   48
      Top             =   780
      Visible         =   0   'False
      Width           =   9045
      Begin VB.CommandButton BotaoTransf 
         Caption         =   "Iniciar Transferências Automáticas"
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
         Left            =   4560
         TabIndex        =   66
         Top             =   930
         Width           =   3165
      End
      Begin VB.Frame Frame6 
         Caption         =   "FTP"
         Height          =   3495
         Left            =   105
         TabIndex        =   49
         Top             =   45
         Width           =   3600
         Begin VB.Frame Frame7 
            Caption         =   "Conexão"
            Height          =   1530
            Left            =   180
            TabIndex        =   58
            Top             =   1815
            Width           =   3300
            Begin VB.CommandButton BotaoFTP 
               Caption         =   "Testar"
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
               Left            =   900
               TabIndex        =   60
               Top             =   1065
               Width           =   1620
            End
            Begin VB.Label FTPComando 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   960
               TabIndex        =   63
               Top             =   210
               Width           =   2250
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Comando:"
               BeginProperty Font 
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
               TabIndex        =   62
               Top             =   255
               Width           =   855
            End
            Begin VB.Label FTPStatus 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   945
               TabIndex        =   61
               Top             =   645
               Width           =   2265
            End
            Begin VB.Label Label8 
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
               Left            =   300
               TabIndex        =   59
               Top             =   690
               Width           =   615
            End
         End
         Begin VB.TextBox FTPDiretorio 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1155
            MaxLength       =   255
            TabIndex        =   57
            Top             =   1485
            Width           =   2295
         End
         Begin VB.TextBox FTPPassword 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1155
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   55
            Top             =   1065
            Width           =   2295
         End
         Begin VB.TextBox FTPUsername 
            Height          =   300
            Left            =   1155
            MaxLength       =   50
            TabIndex        =   53
            Top             =   675
            Width           =   2295
         End
         Begin VB.TextBox FTPURL 
            Height          =   300
            Left            =   1155
            MaxLength       =   255
            TabIndex        =   51
            Top             =   285
            Width           =   2295
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Diretório:"
            BeginProperty Font 
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
            TabIndex        =   56
            Top             =   1545
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Senha:"
            BeginProperty Font 
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
            TabIndex        =   54
            Top             =   1125
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Usuário:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   52
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label4 
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
            Left            =   195
            TabIndex        =   50
            Top             =   330
            Width           =   885
         End
      End
      Begin MSMask.MaskEdBox IntervaloTransf 
         Height          =   315
         Left            =   7410
         TabIndex        =   64
         ToolTipText     =   "Espaço deixado entre os cupons"
         Top             =   375
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Intervalo Entre Transferências (min.):"
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
         Left            =   4215
         TabIndex        =   65
         Top             =   420
         Width           =   3180
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3990
      Index           =   2
      Left            =   150
      TabIndex        =   20
      Top             =   765
      Visible         =   0   'False
      Width           =   9045
      Begin VB.Frame Frame2 
         Caption         =   "Truncamento / Arredondamento"
         Height          =   780
         Index           =   2
         Left            =   375
         TabIndex        =   34
         ToolTipText     =   "Comportamento do sistemas em relação aos valores"
         Top             =   3105
         Width           =   4050
         Begin VB.OptionButton Arredondamento 
            Caption         =   "Arredondamento"
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
            Left            =   2085
            TabIndex        =   36
            Top             =   345
            Width           =   1770
         End
         Begin VB.OptionButton Truncamento 
            Caption         =   "Truncamento"
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
            Left            =   225
            TabIndex        =   35
            Top             =   345
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Impressão do Cupom"
         Height          =   780
         Index           =   0
         Left            =   375
         TabIndex        =   31
         ToolTipText     =   "Quando o cupom deve ser impresso"
         Top             =   2250
         Width           =   4050
         Begin VB.OptionButton ImpAposPagto 
            Caption         =   "Após o pagamento"
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
            Left            =   1860
            TabIndex        =   33
            Top             =   360
            Width           =   1905
         End
         Begin VB.OptionButton ImpItemAItem 
            Caption         =   "Item a item"
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
            TabIndex        =   32
            Top             =   360
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin VB.Frame FrameCartao 
         Caption         =   "Alíquotas ICMS / ISS"
         Height          =   2010
         Left            =   4740
         TabIndex        =   39
         ToolTipText     =   "Alíquotas a serem utilizadas pelo ECF"
         Top             =   1890
         Width           =   4050
         Begin VB.CheckBox ISS 
            Height          =   195
            Left            =   2865
            TabIndex        =   43
            Top             =   300
            Width           =   555
         End
         Begin MSMask.MaskEdBox Sigla 
            Height          =   270
            Left            =   840
            TabIndex        =   41
            Top             =   270
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   476
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Aliquota 
            Height          =   270
            Left            =   1560
            TabIndex        =   42
            Top             =   255
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   476
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCartoes 
            Height          =   1560
            Left            =   255
            TabIndex        =   40
            Top             =   300
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   2752
            _Version        =   393216
            Rows            =   7
            Cols            =   3
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CheckBox CupomDescreveFormaPagto 
         Caption         =   "Cupom descreve forma de pagamento"
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
         Left            =   435
         TabIndex        =   21
         ToolTipText     =   "Indica se no cupom deve aparecer a forma de pagamento"
         Top             =   195
         Width           =   3570
      End
      Begin VB.TextBox MensagemCupom 
         Height          =   1410
         Left            =   4770
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   38
         ToolTipText     =   "Mensagem que deve vir impressa no cupom"
         Top             =   360
         Width           =   4050
      End
      Begin VB.CheckBox HorarioVerao 
         Caption         =   "Horário de Verão"
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
         Left            =   465
         TabIndex        =   30
         ToolTipText     =   "Indica se o sistema está ou não no horário de verão"
         Top             =   1860
         Width           =   1815
      End
      Begin MSMask.MaskEdBox EspacoEntreLinhas 
         Height          =   315
         Left            =   2355
         TabIndex        =   23
         ToolTipText     =   "Espaço deixado entre as linhas no cupom"
         Top             =   555
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LinhasEntreCupons 
         Height          =   315
         Left            =   2370
         TabIndex        =   26
         ToolTipText     =   "Espaço deixado entre os cupons"
         Top             =   990
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox SimboloMoeda 
         Height          =   315
         Left            =   2370
         TabIndex        =   29
         ToolTipText     =   "Símbolo da moeda utilizada"
         Top             =   1440
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "linhas"
         BeginProperty Font 
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
         Left            =   2880
         TabIndex        =   27
         ToolTipText     =   "Espaço deixado entre os cupons"
         Top             =   1050
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dots"
         BeginProperty Font 
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
         Left            =   2880
         TabIndex        =   24
         ToolTipText     =   "Espaço deixado entre as linhas no cupom"
         Top             =   615
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mensagem no Cupom:"
         BeginProperty Font 
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
         Left            =   4770
         TabIndex        =   37
         ToolTipText     =   "Mensagem que deve vir impressa no cupom"
         Top             =   120
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Símbolo de Moeda:"
         BeginProperty Font 
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
         TabIndex        =   28
         ToolTipText     =   "Símbolo da moeda utilizada"
         Top             =   1485
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Espaço entre linhas:"
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
         Left            =   435
         TabIndex        =   22
         ToolTipText     =   "Espaço deixado entre as linhas no cupom"
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Linhas entre cupons:"
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
         Left            =   435
         TabIndex        =   25
         ToolTipText     =   "Espaço deixado entre os cupons"
         Top             =   1050
         Width           =   1800
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4695
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3975
      Index           =   1
      Left            =   210
      TabIndex        =   1
      Top             =   765
      Width           =   9075
      Begin VB.Frame Frame2 
         Caption         =   "Operador e Vendedor"
         Height          =   780
         Index           =   1
         Left            =   4560
         TabIndex        =   8
         ToolTipText     =   "Indica se o Operador e o Vendedor são a mesma pessoa"
         Top             =   1080
         Width           =   4230
         Begin VB.OptionButton OpVendIguais 
            Caption         =   "Iguais"
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
            Left            =   555
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton OpVendDistintos 
            Caption         =   "Distintos"
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
            Left            =   2385
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Caixa Central"
         Height          =   855
         Left            =   285
         TabIndex        =   13
         ToolTipText     =   "Conta contábil do Caixa central dessa filial da empresa"
         Top             =   2655
         Width           =   4050
         Begin MSMask.MaskEdBox MaskContaCentral 
            Height          =   300
            Left            =   1830
            TabIndex        =   15
            Top             =   390
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   529
            _Version        =   393216
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Conta Contábil :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   435
            Width           =   1380
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Operações de Caixa"
         Height          =   915
         Left            =   315
         TabIndex        =   11
         ToolTipText     =   "Indica se o genrente necessita autorizar as operações de Caixa"
         Top             =   1080
         Width           =   4050
         Begin VB.CheckBox GerenteAutoriza 
            Caption         =   "Necessita da autorização do Gerente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   240
            TabIndex        =   12
            Top             =   300
            Width           =   3645
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tela de Venda"
         Height          =   1605
         Left            =   4575
         TabIndex        =   16
         Top             =   1920
         Width           =   4230
         Begin VB.CheckBox AbreAposFecha 
            Caption         =   "Abre cupom após o fechamento de venda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   240
            TabIndex        =   47
            Top             =   240
            Width           =   3885
         End
         Begin VB.OptionButton PoucosProdutos 
            Caption         =   "Com Teclado"
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
            Left            =   420
            TabIndex        =   18
            ToolTipText     =   "Indica qual é a tela de venda utilizada nessa filial da empresa."
            Top             =   1140
            Width           =   1470
         End
         Begin VB.OptionButton MuitosProdutos 
            Caption         =   "Sem Teclado"
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
            Left            =   2280
            TabIndex        =   19
            ToolTipText     =   "Indica qual é a tela de venda utilizada nessa filial da empresa."
            Top             =   1155
            Value           =   -1  'True
            Width           =   1440
         End
         Begin VB.CheckBox ObrigaVendedor 
            Caption         =   "Preenchimento Vendedor obrigatório"
            BeginProperty Font 
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
            TabIndex        =   17
            ToolTipText     =   "Indica se o preenchimento do vendedor na tela de venda é obrigatório"
            Top             =   720
            Width           =   3510
         End
      End
      Begin VB.ComboBox TabelaPrecoLoja 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2010
         TabIndex        =   7
         ToolTipText     =   "Tabela de preços padrão dessa filial da empresa para o loja"
         Top             =   195
         Width           =   2280
      End
      Begin MSMask.MaskEdBox NatOpPadrao 
         Height          =   300
         Left            =   2820
         TabIndex        =   3
         ToolTipText     =   "Natureza de Operação padrão de venda do módulo de loja dessa filial empresa"
         Top             =   600
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
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
      Begin MSMask.MaskEdBox NumLimRO 
         Height          =   300
         Left            =   7815
         TabIndex        =   5
         ToolTipText     =   "Número limite de boletos que devem fazer parte de um resumo de operação (bordero de Boletos)"
         Top             =   600
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
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
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
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
         Height          =   210
         Left            =   330
         TabIndex        =   6
         ToolTipText     =   "Tabela de preços padrão dessa filial da empresa para o loja"
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número Limite de Resumo Operação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4590
         TabIndex        =   4
         ToolTipText     =   "Número limite de boletos que devem fazer parte de um resumo de operação (bordero de Boletos)"
         Top             =   645
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.Label LblNatOp 
         AutoSize        =   -1  'True
         Caption         =   "Natureza Operação Padrão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   375
         TabIndex        =   2
         ToolTipText     =   "Natureza de Operação padrão de venda do módulo de loja dessa filial empresa"
         Top             =   660
         Visible         =   0   'False
         Width           =   2385
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8145
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   90
      Width           =   1140
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "LojaConfig.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   615
         Picture         =   "LojaConfig.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4425
      Left            =   75
      TabIndex        =   0
      Top             =   405
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   7805
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configurações"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ECF"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Transferência de Arquivos"
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
Attribute VB_Name = "LojaConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'****************************************************
'ESSA TELA POSSUI UM FRAME ESCONDIDO (FrameTelaVenda)
'****************************************************

'Property Variables:
Dim m_Caption As String
Event Unload()


'Flag de alteração dos campos da tela
Public iAlterado As Integer

'Indica qual dos frames do tab está visível no momento
Dim iFrameAtual As Integer

'Variável que guarda as características do grid da tela
Dim objGridCartoes As AdmGrid

'Variáveis que guardam o valor das colunas do grid
Dim iGrid_Sigla_Col As Integer 'Coluna de Sigla
Dim iGrid_Aliquota_Col As Integer 'Coluna de Alíquota
Dim iGrid_ISS_Col As Integer 'Coluna de ISS

Private Sub BotaoFTP_Click()

Dim lTeste As Long

On Error GoTo Erro_BotaoFTP_Click

    Inet1.AccessType = icUseDefault
    Inet1.URL = FTPURL.Text
    Inet1.UserName = FTPUsername.Text
    Inet1.Password = FTPPassword.Text
    
    FTPStatus.Caption = ""
    
    lTeste = 0
    FTPComando.Caption = "DIR " & FTPDiretorio.Text & "/*.*"
    BotaoFTP.MousePointer = vbHourglass
    Inet1.Execute , "DIR " & FTPDiretorio.Text & "/*.*"
    Do While FTPStatus.Caption <> "Mensagem completada" And lTeste < 1000000
        lTeste = lTeste + 1
        DoEvents
    Loop
    
    BotaoFTP.MousePointer = vbDefault
    
    If FTPStatus.Caption = "Mensagem completada" Then
        Call Rotina_Aviso(vbOKOnly, "CONEXAO_BEM_SUCEDIDA")
    Else
    
        FTPStatus.Caption = ""
        lTeste = 0
        FTPComando.Caption = "MKDIR " & FTPDiretorio.Text
        BotaoFTP.MousePointer = vbHourglass
        Inet1.Execute , "MKDIR " & FTPDiretorio.Text
        Do While FTPStatus.Caption <> "Mensagem completada" And lTeste < 1000000
            lTeste = lTeste + 1
            DoEvents
        Loop
    
        BotaoFTP.MousePointer = vbDefault
    
        If FTPStatus.Caption = "" Then
            Call Rotina_Aviso(vbOKOnly, "CONEXAO_BEM_SUCEDIDA")
        Else
            Call Rotina_Aviso(vbOKOnly, "NAO_CONSEGUIU_ESTABELECER_CONEXAO")
        End If
    End If
    
    Exit Sub
    
Erro_BotaoFTP_Click:

    BotaoFTP.MousePointer = vbDefault

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162408)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoTransf_Click()

Dim lErro As Long
Dim sNomeArqParam As String
Dim objObject As Object

On Error GoTo Erro_BotaoTransf_Click

    If gobjLoja.lIntervaloTrans > 0 Then
    
        'Prepara para chamar rotina batch
        lErro = Sistema_Preparar_Batch(sNomeArqParam)
        If lErro <> SUCESSO Then gError 133522
            
        gobjLoja.sNomeArqParam = sNomeArqParam
            
        Set gobjLoja.colModulo = gcolModulo
            
        Set objObject = gobjLoja
            
        lErro = CF("Rotina_FTP_Recepcao_CC", objObject)
        If lErro <> SUCESSO And lErro <> 133628 Then gError 133523
        
        If lErro <> SUCESSO Then gError 133636
    
    Else
    
        gError 133524
    
    End If
    
    Exit Sub
    
Erro_BotaoTransf_Click:

    Select Case gErr

        Case 133522, 133523

        Case 133524
            Call Rotina_Erro(vbOKOnly, "ERRO_INTERVALO_TRANSF_NAO_FORNECIDO", gErr)
                    
        Case 133636
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_CARREGOU_ROTINA_RECEPCAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162409)


    End Select

    Exit Sub

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    
Dim sTeste As String
Dim iPOS As Integer
Dim sTeste1 As String

On Error GoTo Erro_Inet1_StateChanged

    Select Case State
    
        Case 1
            FTPStatus.Caption = "Pesquisando IP..."
              
        Case icHostResolved
            FTPStatus.Caption = "IP encontrado"
        
        
        Case icReceivingResponse
            FTPStatus.Caption = "Recebendo mensagem..."
            
        
        Case icResponseCompleted
            sTeste1 = " "
            sTeste = ""
            Do While Len(sTeste1) > 0
                sTeste1 = Inet1.GetChunk(1000)
                sTeste = sTeste & sTeste1
            Loop
            If Left(FTPComando.Caption, 3) = "DIR" And Len(sTeste) < 5 Then
                FTPStatus.Caption = "Diretorio inexistente"
            Else
                FTPStatus.Caption = "Mensagem completada"
            End If
        Case icConnecting
            FTPStatus.Caption = "Conectando..."
            
        Case icConnected
            FTPStatus.Caption = "Conectado"
            
        Case icRequesting
            FTPStatus.Caption = "Enviando pedido ao servidor..."
            
        Case icRequestSent
            FTPStatus.Caption = "Pedido enviado ao servidor"
            
        Case icDisconnecting
            FTPStatus.Caption = "Desconectando..."
            
        Case icDisconnected
            FTPStatus.Caption = "Desconectado"
    
        Case icError
            FTPStatus.Caption = "Erro de comunicação"
    
        Case icResponseReceived
            FTPStatus.Caption = "Mensagem recebida...aguarde"
    
    End Select
    
    Exit Sub
    
Erro_Inet1_StateChanged:

    Select Case gErr
    
    End Select
    
    Exit Sub

End Sub

Private Sub MaskContaCentral_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_MaskContaCentral_Validate

    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", MaskContaCentral.Text, sContaFormatada, objPlanoConta, MODULO_LOJA)
    If lErro <> SUCESSO And lErro <> 5700 Then gError 108195
            
    'conta não cadastrada
    If lErro = 5700 Then gError 108196

    Exit Sub

Erro_MaskContaCentral_Validate:

    Cancel = True


    Select Case gErr
    
        Case 108195
        Case 108196
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", MaskContaCentral.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("PlanoConta", objPlanoConta)

            Else
            
            
            End If
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162410)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'A tela abre com o primeiro frame visível
    iFrameAtual = 1

    'Carrega a ComboBox de tabelas de preços.
    lErro = Carrega_TabelaPreco()
    If lErro <> SUCESSO Then gError 80053

    'Inicializa o Grid da tela
    Set objGridCartoes = New AdmGrid

    lErro = Inicializa_GridCartoes(objGridCartoes)
    If lErro <> SUCESSO Then gError 80068
    
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 108193
    
    'Traz as configurações de loja atuais
    lErro = Traz_LojaConfig_Tela(gobjLoja)
    If lErro <> SUCESSO Then gError 80095
    
    'Zera o flag de alterações indicando que não houve nenhuma ainda
    iAlterado = 0

    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    'Sinaliza erro no carregamento da tela
    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 80053, 80068, 80095
        
        Case 108193

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162411)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Mascaras() As Long

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)

    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 108194
    
    MaskContaCentral.Mask = sMascaraConta
    
    Inicializa_Mascaras = SUCESSO

    Exit Function
    
Erro_Inicializa_Mascaras:
    
    Inicializa_Mascaras = gErr
    
    Select Case gErr
    
        Case 108194

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162412)

    End Select
    
    Exit Function

End Function

Function Carrega_TabelaPreco() As Long
'Carrega na combo as Tabelas de Preço existentes

Dim lErro As Long
Dim objCodDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preço
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELAPRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 80054

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na Lista de Tabela de Preços
        TabelaPrecoLoja.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TabelaPrecoLoja.ItemData(TabelaPrecoLoja.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 80054

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162413)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridCartoes(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Sigla")
    objGridInt.colColuna.Add ("Aliquota")
    objGridInt.colColuna.Add ("ISS")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Sigla.Name)
    objGridInt.colCampo.Add (Aliquota.Name)
    objGridInt.colCampo.Add (ISS.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Sigla_Col = 1
    iGrid_Aliquota_Col = 2
    iGrid_ISS_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridCartoes
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ALIQUOTAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    'Indica suceso na Inicialização
    Inicializa_GridCartoes = SUCESSO

    Exit Function

End Function

Function Traz_LojaConfig_Tela(objLojaConfig As ClassLoja) As Long
'Traz os dados de objLojaConfig para tela

On Error GoTo Erro_Traz_LojaConfig_Tela

    'Frame sem título
    CupomDescreveFormaPagto.Value = objLojaConfig.iCupomDescreveFormaPagto
    EspacoEntreLinhas.Text = objLojaConfig.lEspacoEntreLinhas
    LinhasEntreCupons.Text = objLojaConfig.iLinhasEntreCupons
    SimboloMoeda.Text = objLojaConfig.sSimboloMoeda
    HorarioVerao.Value = objLojaConfig.iHorarioVerao
    GerenteAutoriza.Value = objLojaConfig.iGerenteAutoriza
    
    MaskContaCentral.PromptInclude = False
    MaskContaCentral.Text = objLojaConfig.sContaContabil
    MaskContaCentral.PromptInclude = True

    If objLojaConfig.iImprimeItemAItem = MARCADO Then
        ImpItemAItem.Value = True
    Else
        ImpAposPagto.Value = True
    End If
    
    'verifica se utiliza teclado ou não
    If objLojaConfig.iSemTeclado = MARCADO Then
        MuitosProdutos.Value = True
    Else
        PoucosProdutos.Value = True
    End If
    
    'verifica se é nessário o preench. do vendedor
    ObrigaVendedor.Value = objLojaConfig.iVendedorObrigatorio
    
    'verifica se é nessário o preench. do campo
    AbreAposFecha.Value = objLojaConfig.iAbreAposFechamento
    
    If objLojaConfig.sTruncamentoArredondamento = LOJA_TRUNCAMENTO Then
        Truncamento.Value = True
    Else
        Arredondamento.Value = True
    End If

    If objLojaConfig.iOperadorIgualVendedor = MARCADO Then
        OpVendIguais.Value = True
    Else
        OpVendDistintos.Value = True
    End If

    MensagemCupom.Text = objLojaConfig.sMensagemCupom
'    NatOpPadrao.Text = objLojaConfig.sNatOpPadrao
'    NumLimRO.Text = objLojaConfig.sNumLimRO
    
    If objLojaConfig.iTabelaPreco > 0 Then
        TabelaPrecoLoja.Text = objLojaConfig.iTabelaPreco
        Call TabelaPrecoLoja_Validate(False)
    End If
    
    Call Preenche_GridCartoes(objLojaConfig)
    
    FTPURL.Text = objLojaConfig.sFTPURL
    FTPUsername.Text = objLojaConfig.sFTPUserName
    FTPPassword.Text = objLojaConfig.sFTPPassword
    FTPDiretorio.Text = objLojaConfig.sFTPDiretorio
    
    IntervaloTransf.Text = objLojaConfig.lIntervaloTrans
    
    Exit Function

Erro_Traz_LojaConfig_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162414)

    End Select

    Exit Function

End Function

Function Preenche_GridCartoes(objLojaConfig As ClassLoja) As Long
'Função responsável em transmitir os dados do ObjLojaConfig _
 para o preenchimento do GridCartões

Dim iIndice As Integer
Dim objAliquotaICMS As ClassAliquotaICMS

On Error GoTo Erro_Preenche_GridCartoes

    iIndice = 0

    'Inicia o Preenchimento do Grid
    For Each objAliquotaICMS In objLojaConfig.colAliquotaICMS

        iIndice = iIndice + 1

        GridCartoes.TextMatrix(iIndice, iGrid_Aliquota_Col) = Format(objAliquotaICMS.dAliquota, "Percent")
        GridCartoes.TextMatrix(iIndice, iGrid_ISS_Col) = objAliquotaICMS.iISS
        GridCartoes.TextMatrix(iIndice, iGrid_Sigla_Col) = objAliquotaICMS.sSigla

    Next

    Call Grid_Refresh_Checkbox(objGridCartoes)

    objGridCartoes.iLinhasExistentes = iIndice

    Preenche_GridCartoes = SUCESSO

    Exit Function

Erro_Preenche_GridCartoes:

    Preenche_GridCartoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162415)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Private Sub MaskEdBox1_Change()

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.index).Visible = True
        'Torna Frame atual invisivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.index

    End If

End Sub
Public Sub TabelaPrecoLoja_Validate(Cancel As Boolean)
'Função responsável em validar os dados inseridos na Combo

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPrecoLoja As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPrecoLoja_Validate

    'Verifica se foi preenchida a ComboBox TabelaPrecoLoja
    If Len(Trim(TabelaPrecoLoja.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPrecoLoja
    If TabelaPrecoLoja.Text = TabelaPrecoLoja.List(TabelaPrecoLoja.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPrecoLoja, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 80091

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPrecoLoja.iCodigo = iCodigo

        'Tenta ler TabelaPreçoLoja com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPrecoLoja)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 80093

        If lErro <> SUCESSO Then gError 80094 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreçoLoja no BD, coloca no Text da Combo
        TabelaPrecoLoja.Text = CStr(objTabelaPrecoLoja.iCodigo) & SEPARADOR & objTabelaPrecoLoja.sDescricao

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 80095

    Exit Sub

Erro_TabelaPrecoLoja_Validate:

    Cancel = True

    Select Case gErr

        Case 80091, 80093

        Case 80095, 80094
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPrecoLoja.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162416)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'Sigla
            Case iGrid_Sigla_Col
                lErro = Saida_Celula_Sigla(objGridInt)
                If lErro <> SUCESSO Then gError 80069

            'Aliquota
            Case iGrid_Aliquota_Col
                lErro = Saida_Celula_Aliquota(objGridInt)
                If lErro <> SUCESSO Then gError 80070

            'ISS
            Case iGrid_ISS_Col
                lErro = Saida_Celula_ISS(objGridInt)
                If lErro <> SUCESSO Then gError 80071

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 80072

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 80069 To 80071
            'Variavel não definida
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 80072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162417)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Sigla(objGridCartoes As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Sigla

    Set objGridCartoes.objControle = Sigla

    'Se necessário cria uma nova linha no Grid
    If Len(Trim(Sigla.Text)) > 0 Then
    
        'Verifica se já existe a Sigla no Grid
        For iIndice = 1 To objGridCartoes.iLinhasExistentes

            If iIndice <> GridCartoes.Row Then
                If GridCartoes.TextMatrix(iIndice, iGrid_Sigla_Col) = Sigla.Text Then gError 80219
           End If
        Next
        
        'Se for uma nova linha incrementa o contador de linhas existentes
        If GridCartoes.Row > objGridCartoes.iLinhasExistentes Then
            objGridCartoes.iLinhasExistentes = objGridCartoes.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridCartoes)
    If lErro <> SUCESSO Then gError 80073

    Saida_Celula_Sigla = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Sigla:

       Saida_Celula_Sigla = gErr

    Select Case gErr

        Case 80073
            Call Grid_Trata_Erro_Saida_Celula(objGridCartoes)
        
        Case 80219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_EXISTE", gErr, Sigla)
            Call Grid_Trata_Erro_Saida_Celula(objGridCartoes)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162418)

    End Select

End Function

Private Function Saida_Celula_Aliquota(objGridCartoes As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Aliquota

    Set objGridCartoes.objControle = Aliquota

    'Se necessário cria uma nova linha no Grid
    If Len(Trim(Aliquota.Text)) > 0 Then
    
        lErro = Porcentagem_Critica(Aliquota.Text)
        If lErro <> SUCESSO Then gError 80106
        
        Aliquota.Text = Format(Aliquota.Text, "Fixed")
        
        If GridCartoes.Row > objGridCartoes.iLinhasExistentes Then
            objGridCartoes.iLinhasExistentes = objGridCartoes.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridCartoes)
    If lErro <> SUCESSO Then gError 80074

    Saida_Celula_Aliquota = SUCESSO

    Exit Function

Erro_Saida_Celula_Aliquota:

    Saida_Celula_Aliquota = gErr

    Select Case gErr

        Case 80074, 80106
            Call Grid_Trata_Erro_Saida_Celula(objGridCartoes)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162419)

    End Select

End Function

Private Function Saida_Celula_ISS(objGridCartoes As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ISS

    Set objGridCartoes.objControle = ISS

    lErro = Grid_Abandona_Celula(objGridCartoes)
    If lErro <> SUCESSO Then gError 80075

    Saida_Celula_ISS = SUCESSO

    Exit Function

Erro_Saida_Celula_ISS:

    Saida_Celula_ISS = gErr

    Select Case gErr

        Case 80075
            Call Grid_Trata_Erro_Saida_Celula(objGridCartoes)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162420)

    End Select

    Exit Function

End Function
Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a funcao Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 80077

    Call Rotina_Aviso(vbOKOnly, "AVISO_CONFIGURACAO_GRAVADA")

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 80077

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162421)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objLojaConfig As New ClassLoja

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se espaço entre linhas foi preenchido
    If Len(Trim(EspacoEntreLinhas.Text)) = 0 Then gError 80064

    'Verifica se linhas entre cupons foi preenchido
    If Len(Trim(LinhasEntreCupons.Text)) = 0 Then gError 80065

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objLojaConfig)
    If lErro <> SUCESSO Then gError 80066

    lErro = CF("ConfiguraLoja_Gravar", objLojaConfig)
    If lErro <> SUCESSO Then gError 80068

    Call gobjLoja.Inicializa

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 80064
            Call Rotina_Erro(vbOKOnly, "ERRO_ESPACOENTRELINHAS_NAO_PREENCHIDO", gErr)

        Case 80065
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHASENTRECUPONS_NAO_PREENCHIDO", gErr)

        Case 80066, 80068

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162422)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Function Move_Tela_Memoria(objLojaConfig As ClassLoja) As Long
'Move os dados da tela para memoria

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Recolhe os dados da tela para o objLojaConfig
    objLojaConfig.iCupomDescreveFormaPagto = CupomDescreveFormaPagto.Value
    
    objLojaConfig.lEspacoEntreLinhas = StrParaLong(EspacoEntreLinhas.Text)
    objLojaConfig.iLinhasEntreCupons = StrParaInt(LinhasEntreCupons.Text)

    If ImpItemAItem.Value = True Then
        objLojaConfig.iImprimeItemAItem = MARCADO
    Else
        objLojaConfig.iImprimeItemAItem = DESMARCADO
    End If

'    objLojaConfig.sNatOpPadrao = NatOpPadrao.Text
        
'    objLojaConfig.sNumLimRO = NumLimRO.Text

    If Truncamento.Value = True Then
        objLojaConfig.sTruncamentoArredondamento = LOJA_TRUNCAMENTO
    Else
        objLojaConfig.sTruncamentoArredondamento = LOJA_ARREDONDAMENTO
    End If

    If OpVendIguais.Value = True Then
        objLojaConfig.iOperadorIgualVendedor = MARCADO
    Else
        objLojaConfig.iOperadorIgualVendedor = DESMARCADO
    End If
    
    'preenche se é obrigatório o preenchimento do vendedor
    objLojaConfig.iVendedorObrigatorio = ObrigaVendedor.Value
        
    'preenche se é obrigatório o preenchimento do vendedor
    objLojaConfig.iAbreAposFechamento = AbreAposFecha.Value
        
    'verifica se é com ou sem teclado
    If PoucosProdutos.Value = True Then
        objLojaConfig.iSemTeclado = DESMARCADO
    Else
        objLojaConfig.iSemTeclado = MARCADO
    End If

    objLojaConfig.iTabelaPreco = Codigo_Extrai(TabelaPrecoLoja.Text)
    
    objLojaConfig.iHorarioVerao = HorarioVerao.Value
    objLojaConfig.sMensagemCupom = MensagemCupom.Text
    objLojaConfig.sSimboloMoeda = SimboloMoeda.Text
    objLojaConfig.iGerenteAutoriza = GerenteAutoriza.Value
    objLojaConfig.sContaContabil = MaskContaCentral.ClipText
    
    objLojaConfig.sFTPURL = FTPURL.Text
    objLojaConfig.sFTPUserName = FTPUsername.Text
    objLojaConfig.sFTPPassword = FTPPassword.Text
    objLojaConfig.sFTPDiretorio = FTPDiretorio.Text

    objLojaConfig.lIntervaloTrans = StrParaLong(IntervaloTransf.Text)


    'Chamada a função que recolhe os dados do Grid para memoria
    lErro = Move_GridCartao_Memoria(objLojaConfig)
    If lErro <> SUCESSO Then gError 80063

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 80063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162423)

    End Select

    Exit Function

End Function

Private Function Move_GridCartao_Memoria(objLojaConfig As ClassLoja) As Long
'Move os dados do Grid para a memoria

Dim iIndice As Integer
Dim objAliquotaICMS As New ClassAliquotaICMS

On Error GoTo Erro_Move_GridCartao_Memoria

    For iIndice = 1 To objGridCartoes.iLinhasExistentes

        Set objAliquotaICMS = New ClassAliquotaICMS

        'Armazena os dados da Aliquota
        objAliquotaICMS.iFilialEmpresa = giFilialEmpresa
        objAliquotaICMS.sSigla = GridCartoes.TextMatrix(iIndice, iGrid_Sigla_Col)
        objAliquotaICMS.dAliquota = PercentParaDbl(GridCartoes.TextMatrix(iIndice, iGrid_Aliquota_Col))
        objAliquotaICMS.iISS = StrParaInt(GridCartoes.TextMatrix(iIndice, iGrid_ISS_Col))

        objLojaConfig.colAliquotaICMS.Add objAliquotaICMS

    Next

    Move_GridCartao_Memoria = SUCESSO

    Exit Function

Erro_Move_GridCartao_Memoria:

    Move_GridCartao_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162424)

    End Select

    Exit Function

End Function

Private Sub Aliquota_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Arredondamento_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub CupomDescreveFormaPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EspacoEntreLinhas_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub GridCartoes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridCartoes)

End Sub

Private Sub HorarioVerao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ImpAposPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ImpItemAItem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ISS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LinhasEntreCupons_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MensagemCupom_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MuitosProdutos_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Private Sub NatOpPadrao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub

'Private Sub NumLimRO_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub

Private Sub OpVendDistintos_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OpVendIguais_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PoucosProdutos_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sigla_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SimboloMoeda_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabelaPrecoLoja_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub GridCartoes_LeaveCell()

    Call Saida_Celula(objGridCartoes)

End Sub

Private Sub GridCartoes_EnterCell()
    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridCartoes, iAlterado)

End Sub

Private Sub GridCartoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCartoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridCartoes, iAlterado)
    End If

End Sub

Private Sub GridCartoes_GotFocus()

    Call Grid_Recebe_Foco(objGridCartoes)

End Sub

Private Sub GridCartoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCartoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCartoes, iAlterado)
    End If

End Sub

Private Sub GridCartoes_LostFocus()

    Call Grid_Libera_Foco(objGridCartoes)

End Sub



Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objGridCartoes = Nothing
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Configuração"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LojaConfig"
    
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

Private Sub Truncamento_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
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

Private Sub ISS_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCartoes)

End Sub

Private Sub ISS_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCartoes)

End Sub

Private Sub ISS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCartoes.objControle = ISS
    lErro = Grid_Campo_Libera_Foco(objGridCartoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Aliquota_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCartoes)

End Sub

Private Sub Aliquota_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCartoes)

End Sub

Private Sub Aliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCartoes.objControle = Aliquota
    lErro = Grid_Campo_Libera_Foco(objGridCartoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Sigla_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCartoes)

End Sub

Private Sub Sigla_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCartoes)

End Sub

Private Sub Sigla_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCartoes.objControle = Sigla
    lErro = Grid_Campo_Libera_Foco(objGridCartoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub LinhasEntreCupons_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LinhasEntreCupons_Validate

    If Len(Trim(LinhasEntreCupons.Text)) = 0 Then Exit Sub

    'Faz a critica do valor inserido (linhas entre cupons)
    lErro = Valor_Positivo_Critica(LinhasEntreCupons.Text)
    If lErro <> SUCESSO Then gError 80108

    Exit Sub

Erro_LinhasEntreCupons_Validate:

    Cancel = True

    Select Case gErr

        Case 80108
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162425)

    End Select

    Exit Sub

End Sub


Private Sub EspacoEntreLinhas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EspacoEntreLinhas_Validate

    If Len(Trim(EspacoEntreLinhas.Text)) = 0 Then Exit Sub

    'Faz a critica do valor inserido(Espaco entre linhas)
    lErro = Valor_Positivo_Critica(EspacoEntreLinhas.Text)
    If lErro <> SUCESSO Then gError 80109

    Exit Sub

Erro_EspacoEntreLinhas_Validate:

    Cancel = True

    Select Case gErr

        Case 80109
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162426)

    End Select

    Exit Sub

End Sub

Private Sub GridCartoes_RowColChange()

    Call Grid_RowColChange(objGridCartoes)

End Sub

Private Sub GridCartoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCartoes)

End Sub

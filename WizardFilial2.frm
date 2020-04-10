VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmWizardFilial 
   Appearance      =   0  'Flat
   Caption         =   "Configuração"
   ClientHeight    =   5445
   ClientLeft      =   555
   ClientTop       =   915
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WizardFilial2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8415
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   2
      Left            =   15
      TabIndex        =   7
      Tag             =   "2006"
      Top             =   30
      Width           =   8310
      Begin VB.ComboBox EstoqueAno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardFilial2.frx":014A
         Left            =   4350
         List            =   "WizardFilial2.frx":0169
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1635
         Width           =   855
      End
      Begin VB.ComboBox EstoqueMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardFilial2.frx":01A3
         Left            =   1425
         List            =   "WizardFilial2.frx":01CE
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1635
         Width           =   1545
      End
      Begin MSMask.MaskEdBox IntervaloProducao 
         Height          =   315
         Left            =   2430
         TabIndex        =   25
         Top             =   3285
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Intervalo (em dias) :"
         BeginProperty Font 
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
         Left            =   690
         TabIndex        =   24
         Top             =   3330
         Width           =   1710
      End
      Begin VB.Label Label8 
         Caption         =   "Intervalo entre a produção dos insumos e a produção da mercadoria que utiliza os insumos produzidos:"
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
         Index           =   0
         Left            =   405
         TabIndex        =   23
         Top             =   2595
         Width           =   5040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ano:"
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
         Left            =   3855
         TabIndex        =   16
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mês:"
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
         Left            =   930
         TabIndex        =   17
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label8 
         Caption         =   $"WizardFilial2.frx":0237
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
         Index           =   8
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   5040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Estoque"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   195
         TabIndex        =   19
         Top             =   135
         Width           =   2355
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1770
         Index           =   10
         Left            =   5520
         Picture         =   "WizardFilial2.frx":02E0
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   3
      Left            =   30
      TabIndex        =   20
      Tag             =   "2006"
      Top             =   15
      Width           =   8310
      Begin VB.CheckBox AceitaDiferencaNFPC 
         Caption         =   "Aceita diferença no valor unitário e aliquotas ICMS/IPI entre Notas Fiscais e Pedidos de Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   450
         TabIndex        =   21
         Top             =   2670
         Width           =   4950
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Compras"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   195
         TabIndex        =   22
         Top             =   135
         Width           =   2490
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1230
         Index           =   12
         Left            =   5655
         Picture         =   "WizardFilial2.frx":E8EA
         Stretch         =   -1  'True
         Top             =   255
         Width           =   2280
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   4
      Left            =   0
      TabIndex        =   47
      Tag             =   "2006"
      Top             =   0
      Width           =   8310
      Begin VB.Frame Frame3 
         Caption         =   "Tela de Venda"
         Height          =   1440
         Left            =   195
         TabIndex        =   55
         Top             =   3120
         Width           =   5250
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
            Height          =   315
            Left            =   180
            TabIndex        =   58
            ToolTipText     =   "Indica se o preenchimento do vendedor na tela de venda é obrigatório"
            Top             =   465
            Width           =   3495
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
            Left            =   2040
            TabIndex        =   57
            ToolTipText     =   "Indica qual é a tela de venda utilizada nessa filial da empresa."
            Top             =   1020
            Value           =   -1  'True
            Width           =   1440
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
            Left            =   180
            TabIndex        =   56
            ToolTipText     =   "Indica qual é a tela de venda utilizada nessa filial da empresa."
            Top             =   1020
            Width           =   1470
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operador e Vendedor"
         Height          =   780
         Index           =   1
         Left            =   195
         TabIndex        =   52
         ToolTipText     =   "Indica se o Operador e o Vendedor são a mesma pessoa"
         Top             =   965
         Width           =   5280
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
            Left            =   240
            TabIndex        =   54
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
            Left            =   2040
            TabIndex        =   53
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Operações de Caixa"
         Height          =   1035
         Left            =   195
         TabIndex        =   50
         ToolTipText     =   "Indica se o genrente necessita autorizar as operações de Caixa"
         Top             =   1915
         Width           =   5250
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
            TabIndex        =   51
            Top             =   300
            Width           =   3645
         End
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2205
         Left            =   5640
         Picture         =   "WizardFilial2.frx":1008C
         Top             =   135
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "Módulo - Loja"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   195
         TabIndex        =   49
         Top             =   135
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Defina as confgurações gerais do módulo de Loja."
         BeginProperty Font 
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
         Left            =   360
         TabIndex        =   48
         ToolTipText     =   "Mensagem que deve vir impressa no cupom"
         Top             =   600
         Width           =   4290
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   5
      Left            =   0
      TabIndex        =   26
      Tag             =   "2006"
      Top             =   0
      Width           =   8310
      Begin VB.Frame FrameCartao 
         Caption         =   "Alíquotas ICMS / ISS"
         Height          =   2250
         Left            =   4320
         TabIndex        =   41
         ToolTipText     =   "Alíquotas a serem utilizadas pelo ECF"
         Top             =   2400
         Width           =   3810
         Begin VB.CheckBox ISS 
            Height          =   195
            Left            =   2280
            TabIndex        =   42
            Top             =   600
            Width           =   915
         End
         Begin MSMask.MaskEdBox Sigla 
            Height          =   270
            Left            =   1320
            TabIndex        =   43
            Top             =   480
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
            Left            =   1080
            TabIndex        =   44
            Top             =   960
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
            Height          =   1800
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   3175
            _Version        =   393216
            Rows            =   7
            Cols            =   3
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
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
         Left            =   195
         TabIndex        =   33
         ToolTipText     =   "Indica se o sistema está ou não no horário de verão"
         Top             =   3387
         Width           =   1815
      End
      Begin VB.TextBox MensagemCupom 
         Height          =   570
         Left            =   195
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   32
         ToolTipText     =   "Mensagem que deve vir impressa no cupom"
         Top             =   1260
         Width           =   5100
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
         Left            =   195
         TabIndex        =   31
         ToolTipText     =   "Indica se no cupom deve aparecer a forma de pagamento"
         Top             =   2151
         Width           =   3570
      End
      Begin VB.Frame Frame1 
         Caption         =   "Impressão do Cupom"
         Height          =   780
         Index           =   0
         Left            =   195
         TabIndex        =   28
         ToolTipText     =   "Quando o cupom deve ser impresso"
         Top             =   3840
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   360
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin MSMask.MaskEdBox EspacoEntreLinhas 
         Height          =   315
         Left            =   2040
         TabIndex        =   34
         ToolTipText     =   "Espaço deixado entre as linhas no cupom"
         Top             =   2543
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
         Left            =   2040
         TabIndex        =   35
         ToolTipText     =   "Espaço deixado entre os cupons"
         Top             =   2935
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
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
         Left            =   2565
         TabIndex        =   36
         ToolTipText     =   "Espaço deixado entre os cupons"
         Top             =   2995
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Defina as configurações das Emissoras de Cupom Fiscal."
         BeginProperty Font 
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
         Left            =   360
         TabIndex        =   46
         ToolTipText     =   "Mensagem que deve vir impressa no cupom"
         Top             =   600
         Width           =   4860
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
         Index           =   8
         Left            =   195
         TabIndex        =   40
         ToolTipText     =   "Espaço deixado entre os cupons"
         Top             =   2995
         Width           =   1800
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
         Index           =   7
         Left            =   195
         TabIndex        =   39
         ToolTipText     =   "Espaço deixado entre as linhas no cupom"
         Top             =   2603
         Width           =   1800
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
         Index           =   5
         Left            =   195
         TabIndex        =   38
         ToolTipText     =   "Mensagem que deve vir impressa no cupom"
         Top             =   992
         Width           =   1875
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
         Index           =   4
         Left            =   2565
         TabIndex        =   37
         ToolTipText     =   "Espaço deixado entre as linhas no cupom"
         Top             =   2603
         Width           =   375
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   2205
         Left            =   5625
         Picture         =   "WizardFilial2.frx":1606A
         Top             =   135
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "Módulo - Loja"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   195
         TabIndex        =   27
         Top             =   135
         Width           =   1905
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Termino da Instalação"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Tag             =   "3000"
      Top             =   15
      Width           =   8310
      Begin VB.Label Label10 
         Caption         =   "Pressione o botão ""Terminar"" para que suas configurações sejam gravadas."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   780
         TabIndex        =   59
         Top             =   2655
         Width           =   4275
      End
      Begin VB.Label lblStep 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "A Configuração da Filial está encerrada. "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   5
         Left            =   780
         TabIndex        =   15
         Tag             =   "3001"
         Top             =   630
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   3075
         Index           =   5
         Left            =   5655
         Picture         =   "WizardFilial2.frx":1C048
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2430
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4830
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8310
      Begin VB.Label Label14 
         Caption         =   $"WizardFilial2.frx":2422A
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   150
         TabIndex        =   12
         Top             =   3105
         Width           =   7905
      End
      Begin VB.Label Label12 
         Caption         =   "As próximas telas permitirão que você configure o funcionamento do sistema de acordo com as opções escolhidas."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   3000
         TabIndex        =   13
         Top             =   1875
         Width           =   5055
      End
      Begin VB.Label Label11 
         Caption         =   "A Configuração da Filial está sendo iniciada."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   3000
         TabIndex        =   14
         Top             =   375
         Width           =   5055
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2145
         Index           =   0
         Left            =   120
         Picture         =   "WizardFilial2.frx":242FD
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraStep 
      Caption         =   "Frame5"
      Height          =   1815
      Index           =   0
      Left            =   -10000
      TabIndex        =   8
      Top             =   375
      Width           =   2490
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   4875
      Width           =   8415
      Begin VB.CommandButton cmdNav 
         Caption         =   "Terminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   4
         Left            =   7140
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Prosseguir >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   3
         Left            =   5745
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< Voltar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   4620
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   3450
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Ajuda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "100"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   105
         X2              =   8254
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   8254
         Y1              =   30
         Y2              =   30
      End
   End
End
Attribute VB_Name = "frmWizardFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_STEPS = 7

Const MENSAGEM_TERMINO_CONFIG_FILIAL1 = "A Configuração da Filial "
Const MENSAGEM_TERMINO_CONFIG_FILIAL2 = " da Empresa "
Const MENSAGEM_TERMINO_CONFIG_FILIAL3 = " está encerrada."
Const MENSAGEM_INICIO_CONFIG_FILIAL1 = "A Configuração da Filial "
Const MENSAGEM_INICIO_CONFIG_FILIAL2 = " da Empresa "
Const MENSAGEM_INICIO_CONFIG_FILIAL3 = " está sendo iniciada."

Const RES_ERROR_MSG = 30000

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_5 = 5
Const STEP_FINISH = 6

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = "Configuração da Filial "
Const INTRO_KEY = "Tela de Introdução"
Const SHOW_INTRO = "Exibir Introdução"
Const TOPIC_TEXT = "<TOPIC_TEXT>"

'module level vars
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean

Public VBInst       As VBIDE.VBE
Dim mbFinishOK      As Boolean

'DECLARACAO DE VARIAVEIS GLOBAIS
Public iAlterado As Integer
Dim objConfiguraADM1 As ClassConfiguraADM

'Variável que guarda as características do grid da tela
Dim objGridCartoes As AdmGrid

'Variáveis que guardam o valor das colunas do grid
Dim iGrid_Sigla_Col As Integer 'Coluna de Sigla
Dim iGrid_Aliquota_Col As Integer 'Coluna de Alíquota
Dim iGrid_ISS_Col As Integer 'Coluna de ISS

Private Function LJ_Filial_Gravar_Registro() As Long

Dim lErro As Long
Dim objAliquotaICMS As ClassAliquotaICMS
Dim iIndice As Integer
Dim colConfig As New ColLojaConfig

On Error GoTo Erro_LJ_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_LOJA)

    If lErro = SUCESSO Then
    
        'verifica se os campos obrigatórios estão preenchidos
        If Len(Trim(LinhasEntreCupons.ClipText)) = 0 Then gError 109381
        
        If Len(Trim(EspacoEntreLinhas.ClipText)) = 0 Then gError 109382
        
        'preenche o gobjLoja com os dados da tela
        If OpVendIguais.Value = True Then
            gobjLoja.iOperadorIgualVendedor = MARCADO
        Else
            gobjLoja.iOperadorIgualVendedor = DESMARCADO
        End If
        
        'verifica se é necessária autorização de gerente
        gobjLoja.iGerenteAutoriza = GerenteAutoriza.Value
        
        'verifica se é necessário o preenchimento do nome do vendedor
        gobjLoja.iVendedorObrigatorio = ObrigaVendedor.Value
        
        'verifica se é com muitos produtos (sem teclado) ou com poucos produtos( com teclado)
        If MuitosProdutos.Value = True Then
            gobjLoja.iTelaVendaMP = MARCADO
        Else
            gobjLoja.iTelaVendaMP = DESMARCADO
        End If
        
        '-----------------------------------------------------------------------
        'FIM DO CARREGAMENTO DA PARTE QUE NÃO É REFERENTE AO ECF
        '-----------------------------------------------------------------------
        
        gobjLoja.sMensagemCupom = MensagemCupom.Text
        gobjLoja.iCupomDescreveFormaPagto = CupomDescreveFormaPagto.Value
        gobjLoja.iLinhasEntreCupons = StrParaInt(LinhasEntreCupons.Text)
        gobjLoja.lEspacoEntreLinhas = StrParaLong(EspacoEntreLinhas.Text)
        gobjLoja.iHorarioVerao = HorarioVerao.Value
        
        'verifica se É ITEM A ITEM ou após o pagamento
        If ImpItemAItem.Value = True Then
            gobjLoja.iImprimeItemAItem = MARCADO
        Else
            gobjLoja.iImprimeItemAItem = DESMARCADO
        End If
        
        'carrega a coleção de alíquotas de ICMS/ISS
        For iIndice = 1 To objGridCartoes.iLinhasExistentes
    
            Set objAliquotaICMS = New ClassAliquotaICMS
    
            'Armazena os dados da Aliquota
            objAliquotaICMS.iFilialEmpresa = giFilialEmpresa
            objAliquotaICMS.sSigla = GridCartoes.TextMatrix(iIndice, iGrid_Sigla_Col)
            objAliquotaICMS.dAliquota = PercentParaDbl(GridCartoes.TextMatrix(iIndice, iGrid_Aliquota_Col))
            objAliquotaICMS.iISS = StrParaInt(GridCartoes.TextMatrix(iIndice, iGrid_ISS_Col))
    
            gobjLoja.colAliquotaICMS.Add objAliquotaICMS
    
        Next
        
        'preencho a coleção de configuração na parte que se refere à filial
        lErro = CF("ConfiguraLoja_MoverCampos_ColLojaConfig_Filial", gobjLoja, colConfig)
        If lErro <> SUCESSO Then gError 109389
    
        'chamo a função que grava a configuração
        lErro = gobjLoja.Gravar_Trans(gobjLoja, colConfig)
        If lErro <> SUCESSO Then gError 109390

    End If

    LJ_Filial_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_LJ_Filial_Gravar_Registro:
    
    LJ_Filial_Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 109382
            Call Rotina_Erro(vbOKOnly, "ERRO_ESPACOENTRELINHAS_NAO_PREENCHIDO", gErr)

        Case 109381
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHASENTRECUPONS_NAO_PREENCHIDO", gErr)
        
        Case 109389, 109390
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175905)

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
                If GridCartoes.TextMatrix(iIndice, iGrid_Sigla_Col) = Sigla.Text Then gError 109372
           End If
        Next
        
        'Se for uma nova linha incrementa o contador de linhas existentes
        If GridCartoes.Row > objGridCartoes.iLinhasExistentes Then
            objGridCartoes.iLinhasExistentes = objGridCartoes.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridCartoes)
    If lErro <> SUCESSO Then gError 109371

    Saida_Celula_Sigla = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Sigla:

       Saida_Celula_Sigla = gErr

    Select Case gErr

        Case 109371
            Call Grid_Trata_Erro_Saida_Celula(objGridCartoes)
        
        Case 109372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_EXISTE", gErr, Sigla)
            Call Grid_Trata_Erro_Saida_Celula(objGridCartoes)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175906)

    End Select

End Function

Private Function Saida_Celula_Aliquota(objGridCartoes As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Aliquota

    Set objGridCartoes.objControle = Aliquota

    'Se necessário cria uma nova linha no Grid
    If Len(Trim(Aliquota.Text)) > 0 Then
    
        lErro = Porcentagem_Critica(Aliquota.Text)
        If lErro <> SUCESSO Then gError 109373
        
        Aliquota.Text = Format(Aliquota.Text, "Fixed")
        
        If GridCartoes.Row > objGridCartoes.iLinhasExistentes Then
            objGridCartoes.iLinhasExistentes = objGridCartoes.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridCartoes)
    If lErro <> SUCESSO Then gError 109374

    Saida_Celula_Aliquota = SUCESSO

    Exit Function

Erro_Saida_Celula_Aliquota:

    Saida_Celula_Aliquota = gErr

    Select Case gErr

        Case 109374, 109373
            Call Grid_Trata_Erro_Saida_Celula(objGridCartoes)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175907)

    End Select

End Function

Private Function Saida_Celula_ISS(objGridCartoes As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ISS

    Set objGridCartoes.objControle = ISS

    lErro = Grid_Abandona_Celula(objGridCartoes)
    If lErro <> SUCESSO Then gError 109375

    Saida_Celula_ISS = SUCESSO

    Exit Function

Erro_Saida_Celula_ISS:

    Saida_Celula_ISS = gErr

    Select Case gErr

        Case 109375
            Call Grid_Trata_Erro_Saida_Celula(objGridCartoes)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175908)

    End Select

    Exit Function

End Function

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
                If lErro <> SUCESSO Then gError 109367

            'Aliquota
            Case iGrid_Aliquota_Col
                lErro = Saida_Celula_Aliquota(objGridInt)
                If lErro <> SUCESSO Then gError 109368

            'ISS
            Case iGrid_ISS_Col
                lErro = Saida_Celula_ISS(objGridInt)
                If lErro <> SUCESSO Then gError 109369

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 109370

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 109367 To 109369
            'Variavel não definida
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 109370

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175909)

    End Select

    Exit Function

End Function

Private Sub ISS_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

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

Private Sub Aliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Aliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCartoes.objControle = Aliquota
    lErro = Grid_Campo_Libera_Foco(objGridCartoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Sigla_Change()
    iAlterado = REGISTRO_ALTERADO
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

Private Sub GridCartoes_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCartoes)
End Sub

Private Sub GridCartoes_RowColChange()
    Call Grid_RowColChange(objGridCartoes)
End Sub

Private Sub GridCartoes_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridCartoes)
End Sub

Private Sub GridCartoes_LeaveCell()
    Call Saida_Celula(objGridCartoes)
End Sub

Private Sub GridCartoes_EnterCell()
    Call Grid_Entrada_Celula(objGridCartoes, iAlterado)
End Sub

Private Sub GridCartoes_GotFocus()
    Call Grid_Recebe_Foco(objGridCartoes)
End Sub

Private Sub GridCartoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCartoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridCartoes, iAlterado)
    End If

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

Private Sub LJ_Inicializa_Config()

Dim lErro As Long

On Error GoTo Erro_LJ_Inicializa_Config

    lErro = Valida_Step(MODULO_LOJA)
    
    If lErro = SUCESSO Then
    
        'Inicializa o Grid da tela
        Set objGridCartoes = New AdmGrid
    
        lErro = Inicializa_GridCartoes(objGridCartoes)
        If lErro <> SUCESSO Then gError 109364

        'Frame sem título
        CupomDescreveFormaPagto.Value = gobjLoja.iCupomDescreveFormaPagto
        EspacoEntreLinhas.Text = gobjLoja.lEspacoEntreLinhas
        LinhasEntreCupons.Text = gobjLoja.iLinhasEntreCupons
        HorarioVerao.Value = gobjLoja.iHorarioVerao
        GerenteAutoriza.Value = gobjLoja.iGerenteAutoriza
        
        If gobjLoja.iImprimeItemAItem = MARCADO Then
            ImpItemAItem.Value = True
        Else
            ImpAposPagto.Value = True
        End If
        
        'verifica se utiliza teclado ou não
        If gobjLoja.iSemTeclado = MARCADO Then
            MuitosProdutos.Value = True
        Else
            PoucosProdutos.Value = True
        End If
        
        'verifica se é nessário o preench. do vendedor
        ObrigaVendedor.Value = gobjLoja.iVendedorObrigatorio
        
        If gobjLoja.iOperadorIgualVendedor = MARCADO Then
            OpVendIguais.Value = True
        Else
            OpVendDistintos.Value = True
        End If
    
        MensagemCupom.Text = gobjLoja.sMensagemCupom
        
        Exit Sub
    
    End If
        
    Exit Sub
    
    Exit Sub
    
Erro_LJ_Inicializa_Config:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175910)
    
    End Select
    
    Exit Sub

End Sub

Private Sub cmdNav_Click(Index As Integer)
    
Dim nAltStep As Integer
Dim lHelpTopic As Long
Dim rc As Long
Dim lErro As Long
    
On Error GoTo Erro_cmdNav_Click

    Select Case Index
        Case BTN_HELP
            
            SendKeys "{F1}", True
            
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
LABEL_BTN_BACK:
            nAltStep = mnCurStep - 1
            lErro = SetStep(nAltStep, DIR_BACK)
            If lErro = 44865 Then
                mnCurStep = mnCurStep - 1
                GoTo LABEL_BTN_BACK
            End If
            
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
LABEL_BTN_NEXT:
            nAltStep = mnCurStep + 1
            lErro = SetStep(nAltStep, DIR_NEXT)
            If lErro = 44865 Then
                mnCurStep = mnCurStep + 1
                GoTo LABEL_BTN_NEXT
            End If
            
        Case BTN_FINISH
      
            lErro = Gravar_Registro()
            If lErro <> SUCESSO Then Error 44847
            
            objConfiguraADM1.iConfiguracaoOK = True
            
            Unload Me
            
'            If GetSetting(APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, vbNullString) = vbNullString Then
'                frmConfirm.Show vbModal
'            End If
        
    End Select
    
    Exit Sub
    
Erro_cmdNav_Click:

    Select Case Err

        Case 44847

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175911)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()
    Dim i As Integer
    'init all vars
    mbFinishOK = False
    
    For i = STEP_1 To NUM_STEPS - 1
      fraStep(i).left = -10000
    Next
    
    'Load All string info for Form
    LoadResStrings Me
    
    'Determine 1st Step:
    If GetSetting(APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString) = SHOW_INTRO Then
        Call SetStep(STEP_INTRO, DIR_NEXT)
    Else
        Call SetStep(STEP_1, DIR_NONE)
    End If
    
End Sub

Private Function SetStep(nStep As Integer, nDirection As Integer) As Long
  
Dim lErro As Long, iStep As Integer
  
On Error GoTo Erro_SetSetp
  
    Select Case nStep
    
        Case STEP_INTRO
            
        Case STEP_1
            Me.HelpContextID = IDH_CONFIGURACAO_FILIAL_EMPRESA
            Label11.Caption = MENSAGEM_INICIO_CONFIG_FILIAL1 & gsNomeFilialEmpresa & MENSAGEM_INICIO_CONFIG_FILIAL2 & gsNomeEmpresa & MENSAGEM_INICIO_CONFIG_FILIAL3
      
        Case STEP_2
            lErro = Valida_Step(MODULO_ESTOQUE)
            If lErro <> SUCESSO Then Error 44865
            
            Me.HelpContextID = IDH_CONFIGURACAO_FILIAL_EMPRESA_EST
        
        Case STEP_3
            lErro = Valida_Step(MODULO_COMPRAS)
            If lErro <> SUCESSO Then Error 44865
            
            Me.HelpContextID = IDH_CONFIGURACAO_FILIAL_EMPRESA_COM
            
        Case STEP_4
            lErro = Valida_Step(MODULO_LOJA)
            If lErro <> SUCESSO Then Error 44865
            'Me.HelpContextID = IDH_CONFIGURACAO_FILIAL_EMPRESA_LJ1
            
        Case STEP_5
            lErro = Valida_Step(MODULO_LOJA)
            If lErro <> SUCESSO Then Error 44865
            'Me.HelpContextID = IDH_CONFIGURACAO_FILIAL_EMPRESA_LJ2
            
        Case STEP_FINISH
            lErro = LJ_Parte2_Testa()
            If lErro <> SUCESSO Then Error 41805

            lblStep(5).Caption = MENSAGEM_TERMINO_CONFIG_FILIAL1 & gsNomeFilialEmpresa & MENSAGEM_TERMINO_CONFIG_FILIAL2 & gsNomeEmpresa & MENSAGEM_TERMINO_CONFIG_FILIAL3
            mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).left = 0
    If nStep <> mnCurStep Then
        
        For iStep = STEP_INTRO To STEP_FINISH
        
            If iStep <> nStep Then
                fraStep(iStep).left = -10000
                fraStep(iStep).Enabled = False
            End If
    
        Next
    
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
  
    SetStep = SUCESSO

    Exit Function

Erro_SetSetp:

    SetStep = Err

    Select Case Err
    
        Case 41805, 44865

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175912)

    End Select

    Exit Function
  
End Function

Private Sub LinhasEntreCupons_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LinhasEntreCupons_Validate

    If Len(Trim(LinhasEntreCupons.Text)) = 0 Then Exit Sub

    'Faz a critica do valor inserido (linhas entre cupons)
    lErro = Valor_Positivo_Critica(LinhasEntreCupons.Text)
    If lErro <> SUCESSO Then gError 109379

    Exit Sub

Erro_LinhasEntreCupons_Validate:

    Cancel = True

    Select Case gErr

        Case 109379

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175913)

    End Select

    Exit Sub

End Sub

Private Sub EspacoEntreLinhas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EspacoEntreLinhas_Validate

    If Len(Trim(EspacoEntreLinhas.ClipText)) = 0 Then Exit Sub

    'Faz a critica do valor inserido(Espaco entre linhas)
    lErro = Valor_Positivo_Critica(EspacoEntreLinhas.Text)
    If lErro <> SUCESSO Then gError 109378

    Exit Sub

Erro_EspacoEntreLinhas_Validate:

    Cancel = True

    Select Case gErr

        Case 109378

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175914)

    End Select

    Exit Sub

End Sub

Private Function LJ_Parte2_Testa() As Long

Dim lErro As Long

On Error GoTo Erro_LJ_Parte2_Testa

    lErro = Valida_Step(MODULO_LOJA)
    If lErro = SUCESSO Then
    
        If Len(Trim(EspacoEntreLinhas.ClipText)) = 0 Then gError 109376
        
        If Len(Trim(LinhasEntreCupons.ClipText)) = 0 Then gError 109377
        
    End If

    LJ_Parte2_Testa = SUCESSO
    
    Exit Function
    
Erro_LJ_Parte2_Testa:
    
    LJ_Parte2_Testa = gErr
    
    Select Case gErr
    
        Case 109376
            Call Rotina_Erro(vbOKOnly, "ERRO_ESPACOENTRELINHAS_NAO_PREENCHIDO", gErr)

        Case 109377
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHASENTRECUPONS_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175915)
            
    End Select
    
    Exit Function

End Function


Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = STEP_1 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

    Me.Caption = FRM_TITLE & gsNomeFilialEmpresa & " da Empresa " & gsNomeEmpresa
'    Me.Caption = FRM_TITLE & " - " & LoadResString(fraStep(nStep).Tag)

End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error message
    sTmp = LoadResString(RES_ERROR_MSG)
    'get the specific message
    sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
    'see if we need to save the settings
'    If chkSaveSettings(0).Value = vbChecked Then
      
'        SaveSetting APP_CATEGORY, WIZARD_NAME, "OptionName", Option Value
      
'    End If
    Set objConfiguraADM1 = Nothing
    
''    If mbHelpStarted Then rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
End Sub

Private Function Gravar_Registro() As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lTransacaoDic As Long
Dim lConexao As Long

On Error GoTo Erro_Gravar_Registro
    
    iAlterado = 0
    
    lConexao = GL_lConexaoDic
    
    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirDic
    If lTransacaoDic = 0 Then gError 44961
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 44867
    
    lErro = CTB_Exercicio_Gravar_Registro()
    If lErro <> SUCESSO Then gError 44868
    
    lErro = CR_Filial_Gravar_Registro()
    If lErro <> SUCESSO Then gError 41930
    
    lErro = EST_Filial_Gravar_Registro()
    If lErro <> SUCESSO Then gError 41931
    
    lErro = FAT_Filial_Gravar_Registro()
    If lErro <> SUCESSO Then gError 41932

    lErro = COM_Filial_Gravar_Registro()
    If lErro <> SUCESSO Then gError 74934
    
    lErro = LJ_Filial_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109380
    
    lErro = CF("ModuloFilEmp_Atualiza_Configurado", glEmpresa, giFilialEmpresa, objConfiguraADM1.colModulosConfigurar)
    If lErro <> SUCESSO Then gError 44957
    
    '######################################################################################
    'Inserido por Wagner
    'Após as atualizações do Config pega todos registros da filial MATRIZ
    'e replica para a filial criada se ele ainda não existir (Respeita o que já foi criado)
    lErro = CF("Modulos_Filial_Gravar_Registro", giFilialEmpresa)
    If lErro <> SUCESSO Then gError 140419
    '######################################################################################
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 44869
    
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 44962
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = gErr
    
    Select Case gErr

        Case 44867
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 44868, 44957, 44961, 44962, 41930, 41931, 41932, 74934, 109380, 140419

        Case 44869
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175916)

    End Select

    If gErr <> 44962 Then Call Transacao_Rollback
    Call Transacao_RollbackDic

    Exit Function
    
End Function

Private Function Valida_Step(sModulo As String) As Long

Dim vModulo As Variant

    For Each vModulo In objConfiguraADM1.colModulosConfigurar

        If sModulo = vModulo Then
            Valida_Step = SUCESSO
            Exit Function
        End If
        
    Next
    
    Valida_Step = 44863

End Function

Function Trata_Parametros(objConfiguraADM As ClassConfiguraADM) As Long

On Error GoTo Erro_Trata_Parametros

    Set objConfiguraADM1 = objConfiguraADM
    Call LJ_Inicializa_Config
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175917)
    
    End Select
    
    Exit Function

End Function

Private Function CTB_Exercicio_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_CTB_Exercicio_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTABILIDADE)

    If lErro = SUCESSO Then
        
        lErro = CF("Exercicio_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 44866
        
    End If
    
    CTB_Exercicio_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CTB_Exercicio_Gravar_Registro:
    
    CTB_Exercicio_Gravar_Registro = Err
    
    Select Case Err

        Case 44866

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175918)

    End Select

    Exit Function
    
End Function

Private Function EST_Filial_Gravar_Registro() As Long

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim sIntervaloProducao As String

On Error GoTo Erro_EST_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_ESTOQUE)

    If lErro = SUCESSO Then
        
        If Len(Trim(IntervaloProducao.Text)) > 0 Then
            sIntervaloProducao = IntervaloProducao.Text
        Else
            sIntervaloProducao = "0"
        End If
        
        lErro = CF("EST_Instalacao_Filial", giFilialEmpresa, sIntervaloProducao)
        If lErro <> SUCESSO Then Error 41934
        
        If Len(Trim(EstoqueAno.Text)) = 0 Or EstoqueMes.ListIndex = -1 Then Error 32337
        
        objEstoqueMes.iFilialEmpresa = giFilialEmpresa
        objEstoqueMes.iAno = StrParaInt(EstoqueAno.Text)
        objEstoqueMes.iMes = EstoqueMes.ItemData(EstoqueMes.ListIndex)
        
        lErro = CF("EstoqueMes_Insere", objEstoqueMes)
        If lErro <> SUCESSO Then Error 44969
        
    End If
    
    EST_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_EST_Filial_Gravar_Registro:
    
    EST_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 32337
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_MESANO_ESTOQUE", Err)
        
        Case 41934, 44969

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175919)

    End Select

    Exit Function

End Function

Private Function CR_Filial_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_CR_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTASARECEBER)

    If lErro = SUCESSO Then
        
        lErro = CF("CR_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41933
        
    End If
    
    CR_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CR_Filial_Gravar_Registro:
    
    CR_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175920)

    End Select

    Exit Function
    
End Function

Private Function FAT_Filial_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_FAT_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_FATURAMENTO)

    If lErro = SUCESSO Then
        
        lErro = CF("FAT_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41935
        
    End If
    
    FAT_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_FAT_Filial_Gravar_Registro:
    
    FAT_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41935

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175921)

    End Select

    Exit Function
    
End Function
Private Function COM_Filial_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection
Dim sNFDiferentePC As String

On Error GoTo Erro_COM_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_COMPRAS)

    If lErro = SUCESSO Then
        
        If AceitaDiferencaNFPC.Value = vbChecked Then
        
            sNFDiferentePC = "1"
            
        Else
            sNFDiferentePC = "0"
            
        End If
        
        lErro = COM_Instalacao_Filial(giFilialEmpresa, sNFDiferentePC)
        If lErro <> SUCESSO Then Error 41935
        
    End If
    
    COM_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_COM_Filial_Gravar_Registro:
    
    COM_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41935

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175922)

    End Select

    Exit Function
    
End Function

'???? Subir para RotinasCOM/ClassGravaCOM
Function COM_Instalacao_Filial(iFilialEmpresa As Integer, sDiferenteNFPC As String) As Long
'faz as inicializacoes necessarias à criacao de uma nova filial especificas do modulo

Dim lErro As Long, lComando As Long
Dim lComando2 As Long
Dim sTipo As String
Dim sDescricao As String
Dim iTipo As Integer
Dim sCodigo As String, sConteudo As String

On Error GoTo Erro_COM_Instalacao_Filial

    'a matriz já vem pré-inicializada
    If iFilialEmpresa <> FILIAL_MATRIZ Then
        
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 74890
            
        lErro = Comando_Executar(lComando, "INSERT INTO ComprasConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", "NUM_PROXIMO_CODIGO_COTACAO", iFilialEmpresa, "Número automático da próxima cotação. Depende de FilialEmpresa.", 2, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 74891
        
        lErro = Comando_Executar(lComando, "INSERT INTO ComprasConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", "NUM_PROXIMO_CODIGO_PC", iFilialEmpresa, "Código do próximo Pedido de Compras. Depende de FilialEmpresa.", 2, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 74892
                
        lErro = Comando_Executar(lComando, "INSERT INTO ComprasConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", "NUM_PROX_COMPRADOR", iFilialEmpresa, "Código automático do próximo Comprador da FilialEmpresa.", 0, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 74893
                
        lErro = Comando_Executar(lComando, "INSERT INTO ComprasConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", "NUM_PROXIMO_CODIGO_RC", iFilialEmpresa, "Código da próxima Requisição de Compras. Depende de FilialEmpresa.", 0, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 74894
                
        lErro = Comando_Executar(lComando, "INSERT INTO ComprasConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", "NUM_PROXIMO_CODIGO_RC_MODELO", iFilialEmpresa, "Código da próxima Requisição Modelo. Depende de FilialEmpresa.", 0, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 74895
                
        lErro = Comando_Executar(lComando, "INSERT INTO ComprasConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", "NUM_PROXIMO_CODIGO_CONCORRENCIA", iFilialEmpresa, "Número automático da próxima concorrência. Depende de FilialEmpresa.", 0, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 74896
                
        lErro = Comando_Executar(lComando, "INSERT INTO ComprasConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", "NFISCAL_DIFERENTE_PED_COMPRA", iFilialEmpresa, "0 -> Ñ aceita diferença valor unitário nem aliquota ICM/IPI exceto se aliquotas ñ estiverem preenchidas  1 -> aceita diferenças", 0, sDiferenteNFPC)
        If lErro <> AD_SQL_SUCESSO Then gError 74897
                
        Call Comando_Fechar(lComando)
    
    'se é filial matriz atualiza os dados da configuracao de filial
    Else
    
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 74935

        lComando2 = Comando_Abrir()
        If lComando2 = 0 Then gError 74936
        
        sCodigo = String(STRING_COMCONFIG_CODIGO, 0)
        sDescricao = String(STRING_COMCONFIG_DESCRICAO, 0)
        sConteudo = String(STRING_CONTEUDO, 0)

        lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, FilialEmpresa,Descricao,Tipo,Conteudo FROM ComprasConfig WHERE Codigo=? AND FilialEmpresa=?", 0, sCodigo, iFilialEmpresa, sDescricao, iTipo, sConteudo, "NFISCAL_DIFERENTE_PED_COMPRA", FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 74937

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76005

        lErro = Comando_ExecutarPos(lComando2, "UPDATE ComprasConfig SET Codigo=?,FilialEmpresa=?,Descricao=?,Tipo=?,Conteudo=?", lComando, "NFISCAL_DIFERENTE_PED_COMPRA", FILIAL_MATRIZ, "0 -> Ñ aceita diferença valor unitário nem aliquota ICM/IPI exceto se aliquotas ñ estiverem preenchidas  1 -> aceita diferenças", 0, sDiferenteNFPC)
        If lErro <> AD_SQL_SUCESSO Then gError 74897

        Call Comando_Fechar(lComando)
        Call Comando_Fechar(lComando2)
        
    End If
    
    COM_Instalacao_Filial = SUCESSO
     
    Exit Function
    
Erro_COM_Instalacao_Filial:

    COM_Instalacao_Filial = gErr
     
    Select Case gErr
          
        Case 74890, 74935, 74936
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 74891 To 74897
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ARQCONFIG", gErr)
        
        Case 74937, 76005
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRASCONFIG", gErr, "NFISCAL_DIFERENTE_PED_COMPRA")
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175923)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function




Private Sub lblStep_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(lblStep(Index), Source, X, Y)
End Sub

Private Sub lblStep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblStep(Index), Button, Shift, X, Y)
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

Private Sub Label3_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label3(Index), Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3(Index), Button, Shift, X, Y)
End Sub


Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

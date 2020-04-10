VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl SolicitacaoSRVOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9720
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4875
      Index           =   1
      Left            =   105
      TabIndex        =   19
      Top             =   585
      Width           =   9375
      Begin VB.Frame Frame6 
         Caption         =   "Observação"
         Height          =   1215
         Left            =   45
         TabIndex        =   73
         Top             =   3510
         Width           =   9105
         Begin VB.TextBox Obs 
            Height          =   870
            Left            =   135
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   8820
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Previsão de Entrega"
         Height          =   615
         Left            =   45
         TabIndex        =   71
         Top             =   1335
         Width           =   9105
         Begin VB.Frame Frame9 
            Caption         =   "Fator"
            Height          =   435
            Left            =   2400
            TabIndex        =   89
            Top             =   105
            Width           =   3315
            Begin VB.OptionButton OptPrazoUteis 
               Caption         =   "dias úteis"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1755
               TabIndex        =   9
               Top             =   180
               Value           =   -1  'True
               Width           =   1395
            End
            Begin VB.OptionButton OptPrazoCorr 
               Caption         =   "dias corridos"
               BeginProperty Font 
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
               TabIndex        =   8
               Top             =   180
               Width           =   1665
            End
         End
         Begin MSMask.MaskEdBox Prazo 
            Height          =   300
            Left            =   1185
            TabIndex        =   7
            Top             =   225
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEnt 
            Height          =   300
            Left            =   7590
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   300
            Left            =   6630
            TabIndex        =   10
            ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrerá."
            Top             =   240
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "dias"
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
            Left            =   1920
            TabIndex        =   75
            Top             =   270
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Prazo:"
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
            Left            =   570
            TabIndex        =   74
            Top             =   255
            Width           =   555
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   6075
            TabIndex        =   72
            Top             =   270
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Outros"
         Height          =   1455
         Left            =   45
         TabIndex        =   42
         Top             =   1965
         Width           =   9105
         Begin VB.ComboBox Fase 
            Height          =   315
            Left            =   5580
            TabIndex        =   17
            Top             =   1020
            Width           =   3330
         End
         Begin VB.ComboBox Tipo 
            Height          =   315
            Left            =   1215
            TabIndex        =   16
            Top             =   1020
            Width           =   3330
         End
         Begin VB.TextBox ClienteBenef 
            Height          =   315
            Left            =   1215
            TabIndex        =   14
            ToolTipText     =   "Digite código, nome reduzido, cgc do cliente ou pressione F3 para consulta."
            Top             =   615
            Width           =   3330
         End
         Begin VB.ComboBox FilialClienteBenef 
            Height          =   315
            Left            =   5580
            TabIndex        =   15
            ToolTipText     =   "Digite o nome ou o código da filial do cliente com quem foi feito o relacionamento."
            Top             =   630
            Width           =   2325
         End
         Begin VB.ComboBox Atendente 
            Height          =   315
            Left            =   1215
            TabIndex        =   12
            ToolTipText     =   "Digite o código, o nome do atendente ou aperte F3 para consulta. Para cadastrar novos tipos, use a tela Campos Genéricos."
            Top             =   210
            Width           =   3330
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   315
            Left            =   5580
            TabIndex        =   13
            Top             =   240
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Fase:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5025
            TabIndex        =   106
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label Label13 
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
            Height          =   195
            Left            =   690
            TabIndex        =   105
            Top             =   1065
            Width           =   450
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
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   5040
            TabIndex        =   69
            Top             =   675
            Width           =   465
         End
         Begin VB.Label LabelClienteBenef 
            AutoSize        =   -1  'True
            Caption         =   "Beneficiário:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   75
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   68
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label LabelAtendente 
            AutoSize        =   -1  'True
            Caption         =   "Atendente:"
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
            Left            =   210
            TabIndex        =   44
            Top             =   270
            Width           =   945
         End
         Begin VB.Label LabelVendedor 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4635
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   43
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Dados do Cliente"
         Height          =   570
         Left            =   60
         TabIndex        =   33
         Top             =   720
         Width           =   9090
         Begin VB.ComboBox FilialCliente 
            Height          =   315
            Left            =   5595
            TabIndex        =   6
            ToolTipText     =   "Digite o nome ou o código da filial do cliente com quem foi feito o relacionamento."
            Top             =   210
            Width           =   2250
         End
         Begin VB.TextBox Cliente 
            Height          =   315
            Left            =   1200
            TabIndex        =   5
            ToolTipText     =   "Digite código, nome reduzido, cgc do cliente ou pressione F3 para consulta."
            Top             =   195
            Width           =   3405
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
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   35
            Top             =   240
            Width           =   660
         End
         Begin VB.Label LabelFilialCliente 
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
            Left            =   5070
            TabIndex        =   34
            Top             =   255
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   585
         Left            =   60
         TabIndex        =   29
         Top             =   90
         Width           =   9090
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2190
            Picture         =   "SolicitacaoSRVOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Pressione esse botão para gerar um código automático para o relacionamento."
            Top             =   180
            Width           =   300
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   4365
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   165
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   3405
            TabIndex        =   2
            ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrerá."
            Top             =   165
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Hora 
            Height          =   315
            Left            =   5610
            TabIndex        =   4
            Top             =   165
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1200
            TabIndex        =   0
            Top             =   180
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7425
            TabIndex        =   46
            Top             =   180
            Width           =   1605
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
            Left            =   6795
            TabIndex        =   45
            Top             =   225
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
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
            Index           =   4
            Left            =   5055
            TabIndex        =   32
            Top             =   225
            Width           =   480
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   3
            Left            =   2805
            TabIndex        =   31
            Top             =   225
            Width           =   480
         End
         Begin VB.Label LabelCodigo 
            Caption         =   "Código:"
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
            Left            =   465
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   225
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   675
      Visible         =   0   'False
      Width           =   9330
      Begin VB.Frame Frame3 
         Caption         =   "Solicitações"
         Height          =   4800
         Left            =   0
         TabIndex        =   47
         Top             =   -15
         Width           =   9300
         Begin MSMask.MaskEdBox DataBaixa 
            Height          =   225
            Left            =   6600
            TabIndex        =   104
            Top             =   2520
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
         Begin VB.TextBox Reparo 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4155
            MaxLength       =   250
            TabIndex        =   99
            Top             =   1470
            Width           =   3000
         End
         Begin VB.TextBox DetColuna 
            Height          =   870
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   77
            Top             =   3495
            Width           =   9105
         End
         Begin VB.CommandButton BotaoContrato 
            Caption         =   "Contrato"
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
            Left            =   7620
            TabIndex        =   65
            Top             =   4410
            Width           =   1605
         End
         Begin VB.CommandButton BotaoGarantia 
            Caption         =   "Garantia"
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
            Left            =   5733
            TabIndex        =   64
            Top             =   4410
            Width           =   1605
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
            Left            =   75
            TabIndex        =   61
            Top             =   4410
            Width           =   1605
         End
         Begin VB.CommandButton BotaoServicos 
            Caption         =   "Serviços"
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
            Left            =   1961
            TabIndex        =   62
            Top             =   4410
            Width           =   1605
         End
         Begin VB.CommandButton BotaoLote 
            Caption         =   "Lote"
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
            Left            =   3847
            TabIndex        =   63
            Top             =   4410
            Width           =   1605
         End
         Begin VB.TextBox Solicitacao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4530
            TabIndex        =   66
            Top             =   720
            Width           =   3000
         End
         Begin VB.ComboBox FilialOP 
            Height          =   315
            Left            =   3690
            TabIndex        =   57
            Top             =   1845
            Width           =   2160
         End
         Begin VB.ComboBox UM 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4635
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   375
            Width           =   855
         End
         Begin VB.TextBox ProdutoDesc 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   270
            Left            =   510
            MaxLength       =   50
            TabIndex        =   52
            Top             =   1410
            Width           =   2600
         End
         Begin VB.TextBox ServicoDesc 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   270
            Left            =   3930
            MaxLength       =   50
            TabIndex        =   51
            Top             =   1440
            Width           =   2600
         End
         Begin VB.ComboBox StatusItem 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "SolicitacaoSRVOcx.ctx":00EA
            Left            =   3900
            List            =   "SolicitacaoSRVOcx.ctx":00EC
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   2460
            Width           =   1920
         End
         Begin MSMask.MaskEdBox DataVenda 
            Height          =   225
            Left            =   2415
            TabIndex        =   49
            Top             =   2700
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
         Begin MSMask.MaskEdBox NumIntDoc 
            Height          =   225
            Left            =   6720
            TabIndex        =   50
            Top             =   1995
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   5505
            TabIndex        =   54
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
         Begin MSMask.MaskEdBox Contrato 
            Height          =   225
            Left            =   4965
            TabIndex        =   55
            Top             =   1065
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Garantia 
            Height          =   225
            Left            =   2880
            TabIndex        =   56
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Lote 
            Height          =   270
            Left            =   2055
            TabIndex        =   58
            Top             =   1845
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Servico 
            Height          =   225
            Left            =   3300
            TabIndex        =   59
            Top             =   495
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   840
            TabIndex        =   60
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2895
            Left            =   75
            TabIndex        =   67
            Top             =   210
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   5106
            _Version        =   393216
         End
         Begin VB.Label LinhaDet 
            Height          =   270
            Left            =   3240
            TabIndex        =   102
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label ColunaDet 
            Height          =   270
            Left            =   855
            TabIndex        =   101
            Top             =   3240
            Width           =   1755
         End
         Begin VB.Label Label12 
            Caption         =   "Linha: "
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
            Left            =   2625
            TabIndex        =   100
            Top             =   3240
            Width           =   780
         End
         Begin VB.Label Label8 
            Caption         =   "Coluna: "
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
            Left            =   105
            TabIndex        =   88
            Top             =   3240
            Width           =   780
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   3
      Left            =   105
      TabIndex        =   70
      Top             =   675
      Visible         =   0   'False
      Width           =   9360
      Begin VB.Frame FrameCRM2 
         Caption         =   "Assunto"
         Height          =   2445
         Left            =   165
         TabIndex        =   85
         Top             =   2325
         Width           =   9045
         Begin VB.TextBox AssuntoCRM 
            Height          =   1800
            Left            =   150
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   87
            Top             =   555
            Width           =   8775
         End
         Begin VB.CheckBox MsgAutoCRM 
            Caption         =   "gerar automaticamente"
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
            Left            =   150
            TabIndex        =   86
            Top             =   270
            Value           =   1  'Checked
            Width           =   2475
         End
      End
      Begin VB.CheckBox GravarCRM 
         Caption         =   "Gravar/alterar contato no CRM"
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
         Left            =   180
         TabIndex        =   84
         Top             =   105
         Width           =   3345
      End
      Begin VB.Frame FrameCRM1 
         Caption         =   "Contato"
         Height          =   1890
         Left            =   180
         TabIndex        =   76
         Top             =   360
         Width           =   9030
         Begin VB.CommandButton BotaoLimparCRM 
            Height          =   315
            Left            =   2340
            Picture         =   "SolicitacaoSRVOcx.ctx":00EE
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Numeração Automática"
            Top             =   210
            Width           =   345
         End
         Begin VB.CheckBox Encerrado 
            Caption         =   "Encerrado"
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
            Left            =   2820
            TabIndex        =   98
            Top             =   255
            Width           =   1215
         End
         Begin VB.Frame FrameFim 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   390
            Left            =   4110
            TabIndex        =   92
            Top             =   225
            Width           =   4155
            Begin MSComCtl2.UpDown UpDownDataFim 
               Height          =   300
               Left            =   1635
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataFim 
               Height          =   300
               Left            =   645
               TabIndex        =   94
               ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrerá."
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox HoraFim 
               Height          =   315
               Left            =   2535
               TabIndex        =   95
               Top             =   0
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "hh:mm:ss"
               Mask            =   "##:##:##"
               PromptChar      =   " "
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
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
               Left            =   1995
               TabIndex        =   97
               Top             =   60
               Width           =   480
            End
            Begin VB.Label Label10 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   90
               TabIndex        =   96
               Top             =   60
               Width           =   480
            End
         End
         Begin VB.ComboBox SatisfacaoCRM 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   1425
            Width           =   7665
         End
         Begin VB.ComboBox MotivoCRM 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   1035
            Width           =   7665
         End
         Begin VB.ComboBox StatusCRM 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   660
            Width           =   7665
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            BeginProperty Font 
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
            TabIndex        =   91
            Top             =   270
            Width           =   645
         End
         Begin VB.Label CodigoCRM 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1140
            TabIndex        =   90
            Top             =   225
            Width           =   1230
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Satisfação:"
            BeginProperty Font 
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
            TabIndex        =   83
            Top             =   1485
            Width           =   960
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   81
            Top             =   1080
            Width           =   645
         End
         Begin VB.Label Label5 
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
            Left            =   420
            TabIndex        =   80
            Top             =   705
            Width           =   615
         End
      End
   End
   Begin VB.CommandButton BotaoCRM 
      Caption         =   "CRM"
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
      Left            =   8055
      TabIndex        =   41
      Top             =   5595
      Width           =   1485
   End
   Begin VB.CommandButton BotaoPedido 
      Caption         =   "Pedidos"
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
      Left            =   5055
      TabIndex        =   39
      Top             =   5595
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton BotaoFatura 
      Caption         =   "Faturamentos"
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
      Left            =   6600
      TabIndex        =   40
      Top             =   5595
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton BotaoAcomp 
      Caption         =   "Acompanhamentos"
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
      Left            =   3270
      TabIndex        =   38
      Top             =   5595
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.CommandButton BotaoOS 
      Caption         =   "Ordens de Serviço"
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
      Left            =   1470
      TabIndex        =   37
      Top             =   5595
      Width           =   1740
   End
   Begin VB.CommandButton BotaoOrcamento 
      Caption         =   "Orçamentos"
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
      Left            =   60
      TabIndex        =   36
      Top             =   5595
      Width           =   1365
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6750
      ScaleHeight     =   450
      ScaleWidth      =   2685
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   15
      Width           =   2745
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "SolicitacaoSRVOcx.ctx":0620
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1650
         Picture         =   "SolicitacaoSRVOcx.ctx":079E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1140
         Picture         =   "SolicitacaoSRVOcx.ctx":0CD0
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   630
         Picture         =   "SolicitacaoSRVOcx.ctx":0E5A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   135
         Picture         =   "SolicitacaoSRVOcx.ctx":0FB4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.CheckBox ImprimeGravacao 
      Caption         =   "Imprimir ao gravar"
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
      Left            =   4515
      TabIndex        =   21
      Top             =   210
      Width           =   1935
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5340
      Left            =   75
      TabIndex        =   28
      Top             =   240
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   9419
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Solicitação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contato"
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
Attribute VB_Name = "SolicitacaoSRVOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'??? implementar criação de registro em crfatconfig ao inserir nova filial para relacionamentoclientes e para atendentes

Private glTipoPadrao As Long
Private glFasePadrao As Long

'Eventos de browser
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoContrato As AdmEvento
Attribute objEventoContrato.VB_VarHelpID = -1
Private WithEvents objEventoGarantia As AdmEvento
Attribute objEventoGarantia.VB_VarHelpID = -1
Private WithEvents objEventoClienteBenef As AdmEvento
Attribute objEventoClienteBenef.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFilialCliAlterada As Integer
Dim iVendedorAlterado As Integer
Dim iClienteBenefAlterado As Integer
Dim iFilialCliBenefAlterada As Integer

Dim iStatus_ListIndex_Padrao As Integer
Dim iMotivo_ListIndex_Padrao As Integer
Dim iSatisfacao_ListIndex_Padrao As Integer

Dim gdtDataEntregaAnt As Date
Dim giLinhaDet As Integer
Dim giColunaDet As Integer

Dim objGridItens As AdmGrid

Dim iGrid_Produto_Col As Integer
Dim iGrid_ProdutoDesc_Col As Integer
Dim iGrid_DataVenda_Col As Integer
Dim iGrid_Servico_Col As Integer
Dim iGrid_Reparo_Col As Integer
Dim iGrid_ServicoDesc_Col As Integer
Dim iGrid_Lote_Col As Integer
Dim iGrid_FilialOP_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Solicitacao_Col As Integer
Dim iGrid_Garantia_Col As Integer
Dim iGrid_Contrato_Col As Integer
Dim iGrid_NumIntDoc_Col As Integer
Dim iGrid_StatusItem_Col As Integer
Dim iGrid_DataBaixa_Col As Integer

Dim giFrameAtual As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Private Function Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giFrameAtual = 1
    
    'Inicializa eventos de browser
    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoServico = New AdmEvento
    Set objEventoLote = New AdmEvento
    Set objEventoContrato = New AdmEvento
    Set objEventoGarantia = New AdmEvento
    Set objEventoClienteBenef = New AdmEvento
    
    Set objGridItens = New AdmGrid
    
    'Carrega Ítens das Combos
    lErro = CargaCombo_StatusItem(StatusItem)
    If lErro <> SUCESSO Then gError 195518
    
    Call Inicializa_Grid_Solicitacao(objGridItens)
    
    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", Atendente)
    If lErro <> SUCESSO Then gError 183253
    
    'Coloca data atual como padrão
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 183615
    
    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 183744
    
    'Inicializa a Máscara de Servico
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Servico)
    If lErro <> SUCESSO Then gError 183745
    
    'carregar tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPOSS, Tipo)
    If lErro <> SUCESSO Then gError 186769

    Tipo.AddItem ""
    Tipo.ItemData(Tipo.NewIndex) = 0
    
    glTipoPadrao = Tipo.ListIndex
    
    'carregar tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_FASESS, Fase)
    If lErro <> SUCESSO Then gError 186769

    Fase.AddItem ""
    Fase.ItemData(Fase.NewIndex) = 0
    
    glFasePadrao = Fase.ListIndex
        
    Call Define_Padrao
    
    Status.Caption = STRING_STATUS_ABERTO
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Function
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 183253, 183615, 183744, 183745, 195518
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183254)
    
    End Select
    
End Function

Public Function CargaCombo_StatusItem(objStatusItem As Object)
'Carga dos itens da combo Situação

Dim lErro As Long
Dim bSelecionaPadrao As Boolean

On Error GoTo Erro_CargaCombo_StatusItem

    bSelecionaPadrao = False

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_SOLICSRV_STATUSITEM, objStatusItem, bSelecionaPadrao, False)
    If lErro <> SUCESSO Then gError 195519

    CargaCombo_StatusItem = SUCESSO

    Exit Function

Erro_CargaCombo_StatusItem:

    CargaCombo_StatusItem = gErr

    Select Case gErr
    
        Case 195519

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195520)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Solicitacao(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Desc. Produto")
    objGridInt.colColuna.Add ("Data Venda")
    objGridInt.colColuna.Add ("Lote/Num.Série")
    objGridInt.colColuna.Add ("Filial O.P.")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Desc. Serviço")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Solicitação")
    objGridInt.colColuna.Add ("Reparo")
    objGridInt.colColuna.Add ("Garantia")
    objGridInt.colColuna.Add ("Manutenção")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Data Baixa")
    objGridInt.colColuna.Add ("")

    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (ProdutoDesc.Name)
    objGridInt.colCampo.Add (DataVenda.Name)
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (Servico.Name)
    objGridInt.colCampo.Add (ServicoDesc.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Solicitacao.Name)
    objGridInt.colCampo.Add (Reparo.Name)
    objGridInt.colCampo.Add (Garantia.Name)
    objGridInt.colCampo.Add (Contrato.Name)
    objGridInt.colCampo.Add (StatusItem.Name)
    objGridInt.colCampo.Add (DataBaixa.Name)
    objGridInt.colCampo.Add (NumIntDoc.Name)


    'Controles que participam do Grid
    iGrid_Produto_Col = 1
    iGrid_ProdutoDesc_Col = 2
    iGrid_DataVenda_Col = 3
    iGrid_Lote_Col = 4
    iGrid_FilialOP_Col = 5
    iGrid_Servico_Col = 6
    iGrid_ServicoDesc_Col = 7
    iGrid_UM_Col = 8
    iGrid_Quantidade_Col = 9
    iGrid_Solicitacao_Col = 10
    iGrid_Reparo_Col = 11
    iGrid_Garantia_Col = 12
    iGrid_Contrato_Col = 13
    iGrid_StatusItem_Col = 14
    iGrid_DataBaixa_Col = 15
    iGrid_NumIntDoc_Col = 16

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_SOLICITACOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    NumIntDoc.Width = 0
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    GridItens.ColWidth(iGrid_NumIntDoc_Col) = 0

    Inicializa_Grid_Solicitacao = SUCESSO

End Function

Public Function Trata_Parametros(Optional ByVal objSolicSRV As ClassSolicSRV) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se recebeu um objeto com dados de um relacionamento
    If Not (objSolicSRV Is Nothing) Then
    
        'Lê e traz os dados do relacionamento para a tela
        lErro = Traz_SolicSRV_Tela(objSolicSRV)
        If lErro <> SUCESSO Then gError 183275
        
    End If
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 183275
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183276)
    
    End Select
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoProduto = Nothing
    Set objEventoServico = Nothing
    Set objEventoLote = Nothing
    Set objEventoContrato = Nothing
    Set objEventoGarantia = Nothing
    Set objEventoClienteBenef = Nothing

    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Sub BotaoLimparCRM_Click()
    CodigoCRM.Caption = ""
End Sub

Private Sub BotaoLote_Click()

Dim colSelecao As New Collection
Dim objRastroLote As New ClassRastreamentoLote
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sSelecao As String
Dim lErro As Long

On Error GoTo Erro_BotaoLote_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 195714

    If Me.ActiveControl Is Lote Then

        objRastroLote.sCodigo = Lote.Text

    Else
    
        objRastroLote.sCodigo = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
    
    End If

    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195715

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 195716

    'Selecao
    colSelecao.Add sProdutoFormatado

    sSelecao = "Produto = ?"

    'Chama tela de Browse de RastreamentoLote
    Call Chama_Tela("RastroLoteLista1", colSelecao, objRastroLote, objEventoLote, sSelecao)

    Exit Sub

Erro_BotaoLote_Click:

    Select Case gErr

        Case 195714
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 195715

        Case 195716
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, GridItens.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195717)

    End Select

    Exit Sub

End Sub

Private Sub FilialCliente_Click()

Dim lErro As Long

On Error GoTo Erro_FilialCliente_Click

    'Faz a validação da filial do cliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 183287
    
    Exit Sub
    
Erro_FilialCliente_Click:

    Select Case gErr

        Case 183287
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183288)

    End Select
    
End Sub

Private Sub GravarCRM_Click()
    Call Trata_GravarCRM
End Sub

Private Sub MsgAutoCRM_Click()
    Call Trata_MsgAutoCRM
End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As ClassRastreamentoLote
Dim iCodigo As Integer

On Error GoTo Erro_objEventoLote_evSelecao

    Set objRastroLote = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    Lote.Text = objRastroLote.sCodigo

    If Not (Me.ActiveControl Is Lote) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col) = Lote.Text

    End If
    
    If objRastroLote.iFilialOP <> 0 Then

        FilialOP.Text = objRastroLote.iFilialOP

        'Tenta selecionar na combo
        lErro = Combo_Seleciona(FilialOP, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195719

        If lErro = SUCESSO Then GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col) = FilialOP.Text

    End If

    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case gErr

        Case 195719

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195718)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 183277
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 183277
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183278)

    End Select

    Exit Sub

End Sub

Public Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se a Data foi digitada
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 183279
    
    Call Trata_Prazo

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 183279

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183280)

    End Select

    Exit Sub

End Sub

Public Sub Hora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Hora, iAlterado)

End Sub

Public Sub Hora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Verifica se a hora foi digitada
    If Len(Trim(Hora.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(Hora.Text)
    If lErro <> SUCESSO Then gError 183281

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 183281

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183282)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cliente_Validate

    'Faz a validação do cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 183285
    
    Exit Sub
    
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 183285
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183286)

    End Select

End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 183283

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 183283

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183284)

    End Select
    
    Exit Sub

End Sub

Private Sub ClienteBenef_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iClienteBenefAlterado = REGISTRO_ALTERADO

    Call ClienteBenef_Preenche

End Sub

Private Sub ClienteBenef_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ClienteBenef_Validate

    'Faz a validação do cliente
    lErro = Valida_ClienteBenef()
    If lErro <> SUCESSO Then gError 210293
    
    Exit Sub
    
Erro_ClienteBenef_Validate:

    Cancel = True

    Select Case gErr

        Case 210293
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210294)

    End Select

End Sub

Private Sub ClienteBenef_Preenche()

Static sNomeReduzidoParteBenef As String
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_ClienteBenef_Preenche
    
    Set objcliente = ClienteBenef
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParteBenef)
    If lErro <> SUCESSO Then gError 210295

    Exit Sub

Erro_ClienteBenef_Preenche:

    Select Case gErr

        Case 210295

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210296)

    End Select
    
    Exit Sub

End Sub

Private Sub FilialCliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iFilialCliAlterada = REGISTRO_ALTERADO
End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FilialCliente_Validate

    'Faz a validação da filial do cliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 183287
    
    Exit Sub
    
Erro_FilialCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 183287
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183288)

    End Select

End Sub

Private Function Valida_FilialCliente() As Long
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer
Dim sCliente As String
Dim objcliente As New ClassCliente

On Error GoTo Erro_Valida_FilialCliente

    'Se a filial de cliente não foi alterada => sai da função
    If iFilialCliAlterada = 0 Then Exit Function
    
    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialCliente.Text)) > 0 Then

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(FilialCliente, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 183289
    
        'Se não encontrou o CÓDIGO
        If lErro = 6730 Then
    
            'Verifica se o cliente foi digitado
            If Len(Trim(Cliente.Text)) = 0 Then gError 183290
    
            sCliente = Cliente.Text
            objFilialCliente.iCodFilial = iCodigo
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 183291
            
            If lErro = 17660 Then
                
                'Lê o Cliente
                objcliente.sNomeReduzido = sCliente
                lErro = CF("Cliente_Le_NomeReduzido", objcliente)
                If lErro <> SUCESSO And lErro <> 12348 Then gError 183292
                
                'Não encontrou Cliente
                If lErro = 12348 Then gError 183293
                
                objFilialCliente.lCodCliente = objcliente.lCodigo
            
                gError 183294
                
            End If
    
            'Coloca na tela a Filial lida
            FilialCliente.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
    
        'Não encontrou a STRING
        ElseIf lErro = 6731 Then
            gError 183295
    
        End If

    End If
    
    iFilialCliAlterada = 0
    
    Valida_FilialCliente = SUCESSO
    
    Exit Function

Erro_Valida_FilialCliente:

    Valida_FilialCliente = gErr

    Select Case gErr

        Case 183289, 183291, 183292, 183293

        Case 183290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 183294
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
            
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 183295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialCliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183296)

    End Select

    Exit Function

End Function

Private Sub FilialClienteBenef_Change()
    iAlterado = REGISTRO_ALTERADO
    iFilialCliBenefAlterada = REGISTRO_ALTERADO
End Sub

Private Sub FilialClienteBenef_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FilialClienteBenef_Validate

    'Faz a validação da filial do cliente
    lErro = Valida_FilialClienteBenef()
    If lErro <> SUCESSO Then gError 210297
    
    Exit Sub
    
Erro_FilialClienteBenef_Validate:

    Cancel = True

    Select Case gErr

        Case 210297
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210298)

    End Select

End Sub

Private Function Valida_FilialClienteBenef() As Long
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer
Dim sCliente As String
Dim objcliente As New ClassCliente

On Error GoTo Erro_Valida_FilialClienteBenef

    'Se a filial de cliente não foi alterada => sai da função
    If iFilialCliBenefAlterada = 0 Then Exit Function
    
    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialClienteBenef.Text)) > 0 Then

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(FilialClienteBenef, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 210299
    
        'Se não encontrou o CÓDIGO
        If lErro = 6730 Then
    
            'Verifica se o cliente foi digitado
            If Len(Trim(ClienteBenef.Text)) = 0 Then gError 210300
    
            sCliente = ClienteBenef.Text
            objFilialCliente.iCodFilial = iCodigo
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 210301
            
            If lErro = 17660 Then
                
                'Lê o Cliente
                objcliente.sNomeReduzido = sCliente
                lErro = CF("Cliente_Le_NomeReduzido", objcliente)
                If lErro <> SUCESSO And lErro <> 12348 Then gError 210302
                
                'Não encontrou Cliente
                If lErro = 12348 Then gError 210303
                
                objFilialCliente.lCodCliente = objcliente.lCodigo
            
                gError 210304
                
            End If
    
            'Coloca na tela a Filial lida
            FilialClienteBenef.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
    
        'Não encontrou a STRING
        ElseIf lErro = 6731 Then
            gError 201305
    
        End If

    End If
    
    iFilialCliBenefAlterada = 0
    
    Valida_FilialClienteBenef = SUCESSO
    
    Exit Function

Erro_Valida_FilialClienteBenef:

    Valida_FilialClienteBenef = gErr

    Select Case gErr

        Case 210299, 210301, 210302, 210303

        Case 210300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 210304
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, ClienteBenef.Text)
            
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 210305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialClienteBenef.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210306)

    End Select

    Exit Function

End Function

Private Sub Atendente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Atendente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Atendente_Validate

    'Valida o atendente selecionado pelo cliente
    lErro = CF("Atendente_Validate", Atendente)
    If lErro <> SUCESSO Then gError 183297
    
    Exit Sub

Erro_Atendente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 183297
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183298)

    End Select

End Sub

Private Sub OptPrazoCorr_Click()
    Call Trata_Prazo
End Sub

Private Sub OptPrazoUteis_Click()
    Call Trata_Prazo
End Sub

Private Sub Prazo_Validate(Cancel As Boolean)
    Call Trata_Prazo
End Sub

Public Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Vendedor, objVendedor)
        If lErro <> SUCESSO Then gError 183300
        
        If objVendedor.iAtivo = DESMARCADO Then gError 183301

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 183300
        
        Case 183301
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183302)
    
    End Select

End Sub

Private Sub LabelCodigo_Click()

Dim objSolicSRV As New ClassSolicSRV
Dim colSelecao As New Collection

    objSolicSRV.lCodigo = StrParaLong(Codigo.Text)
    
    Call Chama_Tela("SolicitacaoSRVLista", colSelecao, objSolicSRV, objEventoCodigo)
    
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objSolicSRV As ClassSolicSRV
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objSolicSRV = obj1
    
    'Traz para a tela o relacionamento com código passado pelo browser
    lErro = Traz_SolicSRV_Tela(objSolicSRV)
    If lErro <> SUCESSO Then gError 183307
        
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 183307
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183308)
    
    End Select

End Sub

Private Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelCliente_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(Cliente.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(Cliente.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = Cliente.Text
        
        sOrdenacao = "Nome Reduzido"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, "", sOrdenacao)

    Exit Sub
    
Erro_LabelCliente_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183312)
    
    End Select
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 183311

    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr
    
        Case 183311
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183313)
    
    End Select

End Sub

Private Sub LabelClienteBenef_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteBenef_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteBenef.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(ClienteBenef.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = ClienteBenef.Text
        
        sOrdenacao = "Nome Reduzido"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoClienteBenef, "", sOrdenacao)

    Exit Sub
    
Erro_LabelClienteBenef_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210310)
    
    End Select
    
End Sub

Private Sub objEventoClienteBenef_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteBenef_evSelecao

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteBenef.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    lErro = Valida_ClienteBenef()
    If lErro <> SUCESSO Then gError 210311

    Me.Show

    Exit Sub

Erro_objEventoClienteBenef_evSelecao:

    Select Case gErr
    
        Case 210311
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210312)
    
    End Select

End Sub

Public Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
On Error GoTo Erro_LabelVendedor_Click
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

    Exit Sub
    
Erro_LabelVendedor_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183314)
    
    End Select

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 183315

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 183316
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto
    
    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
        
Erro_BotaoProdutos_Click:
    
    Select Case gErr
        
        Case 183315
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 183316
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183317)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 183745

    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True

    If Not (Me.ActiveControl Is Produto) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
        'Faz o Tratamento do produto
        lErro = Traz_Produto_Tela()
        If lErro <> SUCESSO Then gError 183746

    End If
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 183745
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""
            
        Case 183746
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183747)

    End Select

    Exit Sub

End Sub

Private Function Traz_Produto_Tela() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_Produto_Tela

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 183321
    
    If lErro = 51381 Then gError 183322

    'Descricao Produto
    GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoDesc_Col) = objProduto.sDescricao

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

    End If

    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 183321

        Case 183322
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183323)

    End Select

    Exit Function

End Function

Public Sub BotaoServicos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sSelecaoSQL As String

On Error GoTo Erro_BotaoServicos_Click

    If Me.ActiveControl Is Servico Then
    
        sProduto1 = Servico.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 183324

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 183325
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto
    
    Set colSelecao = New Collection
    
    colSelecao.Add NATUREZA_PROD_SERVICO
        
    sSelecaoSQL = "Natureza=?"
            
    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico, sSelecaoSQL)

    Exit Sub
        
Erro_BotaoServicos_Click:
    
    Select Case gErr
        
        Case 183324
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 183325
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183326)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 183327

    Servico.PromptInclude = False
    Servico.Text = sProduto
    Servico.PromptInclude = True

    If Not (Me.ActiveControl Is Servico) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col) = Servico.Text
    
        'Faz o Tratamento do produto
        lErro = Traz_Servico_Tela()
        If lErro <> SUCESSO Then gError 183328

    End If
    
    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 183327
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 183328
            GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col) = ""
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183329)

    End Select

    Exit Sub

End Sub

Private Function Traz_Servico_Tela() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_Servico_Tela

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", Servico.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 183330
    
    If lErro = 51381 Then gError 183331

    'Descricao Servico
    GridItens.TextMatrix(GridItens.Row, iGrid_ServicoDesc_Col) = objProduto.sDescricao

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

    End If

    Traz_Servico_Tela = SUCESSO

    Exit Function

Erro_Traz_Servico_Tela:

    Traz_Servico_Tela = gErr

    Select Case gErr

        Case 183330

        Case 183331
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183332)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 183731

    'Limpa a Tela
    Call Limpa_SolicSRV

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 183731

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183732)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objSolicSRV As New ClassSolicSRV
Dim lErro As Long
Dim sAviso As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Se o código não foi preenchido => erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 183715

    'Guarda no obj, código do relacionamento e filial empresa
    'Essas informações são necessárias para excluir o relacionamento
    objSolicSRV.lCodigo = StrParaLong(Codigo.Text)
    objSolicSRV.iFilialEmpresa = giFilialEmpresa

    'Lê o relacionamento com os filtros passados
    lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
    If lErro <> SUCESSO And lErro <> 183258 Then gError 183716
    
    'Se não encontrou => erro
    If lErro <> SUCESSO Then gError 183717
    
    'Pede a confirmação da exclusão da solicitacao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_SOLICITACAOSRV")
    
    If vbMsgRes = vbYes Then

        'Faz a exclusão da Solicitacao
        lErro = CF("SolicitacaoSRV_Exclui", objSolicSRV)
        If lErro <> SUCESSO Then gError 183718
    
        'Limpa a Tela de Orcamento de Venda
        Call Limpa_SolicSRV
        
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 183715
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 183716, 183718

        Case 183717
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183719)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 183729

    'Limpa a Tela
    Call Limpa_SolicSRV
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 183729

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183730)

    End Select

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código de relacionamento para giFilialEmpresa
    lErro = CF("Config_ObterAutomatico", "SRVConfig", "NUM_PROX_SOLICITACAOSRV_1", "SolicitacaoSRV", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 183733
    
    'Exibe o código obtido
    Codigo.PromptInclude = False
    Codigo.Text = lCodigo
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 183733
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183734)

    End Select

End Sub

Private Sub TabStrip1_Click()

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index
       
    End If
    
    Call Trata_AssuntoCRM

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183735)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 183736
    
    Call Trata_Prazo

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 183736

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183737)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 183738
    
    Call Trata_Prazo

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 183738

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183739)

    End Select

    Exit Sub

End Sub

Private Sub Atendente_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub


'*** TRATAMENTO DO EVENTO KEYDOWN  - INÍCIO ***
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call LabelVendedor_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Servico Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is Lote Then
            Call BotaoLote_Click
        ElseIf Me.ActiveControl Is Contrato Then
            Call BotaoContrato_Click
        ElseIf Me.ActiveControl Is Garantia Then
            Call BotaoGarantia_Click
        ElseIf Me.ActiveControl Is ClienteBenef Then
            Call LabelClienteBenef_Click
        End If
    
    End If

End Sub


'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Solicitação de Serviços"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "SolicitacaoSRV"
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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
    RaiseEvent Unload
End Sub

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - INÍCIO ***
Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteBenef_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteBenef, Source, X, Y)
End Sub

Private Sub LabelClienteBenef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteBenef, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialCliente, Source, X, Y)
End Sub

Private Sub LabelFilialCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelAtendente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAtendente, Source, X, Y)
End Sub

Private Sub LabelAtendente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAtendente, Button, Shift, X, Y)
End Sub

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - FIM ***


Private Function Traz_SolicSRV_Tela(ByVal objSolicSRV As ClassSolicSRV) As Long
'Traz pra tela os dados da solicitacao de servico passado como parâmetro

Dim lErro As Long
Dim bCancel As Boolean
Dim iAchou As Integer
Dim iIndice As Integer

On Error GoTo Erro_Traz_SolicSRV_Tela

    'Limpa a tela
    Call Limpa_SolicSRV
    
    Set objSolicSRV.colItens = New Collection
    
    'Lê no BD os dados da solicitacao em questao
    lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
    If lErro <> SUCESSO And lErro <> 183258 Then gError 183266
    
    'Se não encontrou a solicitacao => erro
    If lErro <> SUCESSO Then gError 183267
    
    
    Codigo.PromptInclude = False
    Codigo.Text = objSolicSRV.lCodigo
    Codigo.PromptInclude = True

    If objSolicSRV.dtData <> DATA_NULA Then
        Data.PromptInclude = False
        Data.Text = Format(objSolicSRV.dtData, "dd/mm/yy")
        Data.PromptInclude = True
    End If
    
    If objSolicSRV.dtHora <> 0 Then
        Hora.PromptInclude = False
        Hora.Text = Format(objSolicSRV.dtHora, "hh:mm:ss")
        Hora.PromptInclude = True
    End If
    
    'Se o código do cliente está preenchido
    If objSolicSRV.lCliente <> 0 Then
    
        Call Cliente_Formata(objSolicSRV.lCliente)

        'Se a filial do cliente está preenchida
        If objSolicSRV.iFilial <> 0 Then

            Call Filial_Formata(FilialCliente, objSolicSRV.iFilial)
            
        End If
        
    End If
    
    'Se o código do cliente está preenchido
    If objSolicSRV.lClienteBenef <> 0 Then
    
        Call ClienteBenef_Formata(objSolicSRV.lClienteBenef)

        'Se a filial do cliente está preenchida
        If objSolicSRV.iFilialClienteBenef <> 0 Then

            Call FilialClienteBenef_Formata(FilialClienteBenef, objSolicSRV.iFilialClienteBenef)
            
        End If
        
    End If
    
    If objSolicSRV.iAtendente > 0 Then
        Atendente.Text = objSolicSRV.iAtendente
        Call Atendente_Validate(bCancel)
    End If
    
    If objSolicSRV.iVendedor > 0 Then
        Vendedor.Text = CStr(objSolicSRV.iVendedor)
        Call Vendedor_Validate(bCancel)
    End If
    
    lErro = Carrega_Grid_Itens(objSolicSRV)
    If lErro <> SUCESSO Then gError 183303
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If UCase(GridItens.TextMatrix(iIndice, iGrid_StatusItem_Col)) <> UCase(STRING_BAIXADA) Then
            iAchou = 1
            Exit For
        End If
        
    Next
    
    If iAchou = 0 Then
        Status.Caption = STRING_STATUS_BAIXADO
    Else
        Status.Caption = STRING_STATUS_ABERTO
    End If
    
    Obs.Text = objSolicSRV.sOBS
    
    If objSolicSRV.iPrazoTipo = 0 Then
        OptPrazoUteis.Value = True
    Else
        OptPrazoCorr.Value = True
    End If
    
    If objSolicSRV.iPrazo Then
        Prazo.PromptInclude = False
        Prazo.Text = CStr(objSolicSRV.iPrazo)
        Prazo.PromptInclude = False
    End If
    
    Call DateParaMasked(DataEntrega, objSolicSRV.dtDataEntrega)
    
    lErro = Traz_CRM_Tela(objSolicSRV)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    MsgAutoCRM.Value = vbUnchecked
    Call Trata_MsgAutoCRM
    
    If objSolicSRV.lTipo <> 0 Then
        Call Combo_Seleciona_ItemData(Tipo, objSolicSRV.lTipo)
    Else
        Tipo.ListIndex = -1
    End If
    
    If objSolicSRV.lFase <> 0 Then
        Call Combo_Seleciona_ItemData(Fase, objSolicSRV.lFase)
    Else
        Fase.ListIndex = -1
    End If
    
    iAlterado = 0
    
    Traz_SolicSRV_Tela = SUCESSO

    Exit Function

Erro_Traz_SolicSRV_Tela:

    Traz_SolicSRV_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 183266, 183303
        
        Case 183267
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183268)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_Itens(objSolicSRV As ClassSolicSRV) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sProdutoEnxuto As String
Dim sServicoEnxuto As String
Dim objItemSolicSRV As ClassItensSolicSRV

On Error GoTo Erro_Carrega_Grid_Itens

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For iIndice = 1 To objSolicSRV.colItens.Count
       
        Set objItemSolicSRV = objSolicSRV.colItens(iIndice)
       
        lErro = Mascara_RetornaProdutoEnxuto(objItemSolicSRV.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 183304

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemSolicSRV.sServico, sServicoEnxuto)
        If lErro <> SUCESSO Then gError 183305

        'Mascara o produto enxuto
        Servico.PromptInclude = False
        Servico.Text = sServicoEnxuto
        Servico.PromptInclude = True
        
        GridItens.TextMatrix(iIndice, iGrid_Servico_Col) = Servico.Text
        GridItens.TextMatrix(iIndice, iGrid_ProdutoDesc_Col) = objItemSolicSRV.sProdutoDesc
        If objItemSolicSRV.dtDataVenda <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataVenda_Col) = Format(objItemSolicSRV.dtDataVenda, "dd/mm/yyyy")
        GridItens.TextMatrix(iIndice, iGrid_ServicoDesc_Col) = objItemSolicSRV.sServicoDesc
        GridItens.TextMatrix(iIndice, iGrid_UM_Col) = objItemSolicSRV.sUM
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemSolicSRV.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_Lote_Col) = objItemSolicSRV.sLote
        If objItemSolicSRV.iFilialOP <> 0 Then GridItens.TextMatrix(iIndice, iGrid_FilialOP_Col) = objItemSolicSRV.iFilialOP
        GridItens.TextMatrix(iIndice, iGrid_Solicitacao_Col) = objItemSolicSRV.sSolicitacao
        GridItens.TextMatrix(iIndice, iGrid_Reparo_Col) = objItemSolicSRV.sReparo
        If objItemSolicSRV.lGarantia <> 0 Then GridItens.TextMatrix(iIndice, iGrid_Garantia_Col) = objItemSolicSRV.lGarantia
        GridItens.TextMatrix(iIndice, iGrid_Contrato_Col) = objItemSolicSRV.sContrato
        GridItens.TextMatrix(iIndice, iGrid_NumIntDoc_Col) = objItemSolicSRV.lNumIntDoc
        
        'preenche StatusItem
        For iIndice1 = 0 To StatusItem.ListCount - 1
            If StatusItem.ItemData(iIndice1) = objItemSolicSRV.iStatusItem Then
                StatusItem.ListIndex = iIndice1
                Exit For
            End If
        Next
        
        GridItens.TextMatrix(iIndice, iGrid_StatusItem_Col) = StatusItem.Text
        If objItemSolicSRV.dtDataBaixa <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataBaixa_Col) = Format(objItemSolicSRV.dtDataBaixa, "dd/mm/yyyy")
        
    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = objSolicSRV.colItens.Count

    Carrega_Grid_Itens = SUCESSO

    Exit Function

Erro_Carrega_Grid_Itens:

    Carrega_Grid_Itens = gErr

    Select Case gErr

        Case 183304
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemSolicSRV.sProduto)

        Case 183305
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemSolicSRV.sServico)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183306)

    End Select

    Exit Function

End Function

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente
    
    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
    If lErro <> SUCESSO Then gError 183269

    lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 183270

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

    
    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr
    
        Case 183269, 183270
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183271)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO And lErro <> 17660 Then gError 183272

    If lErro = 17660 Then gError 183273

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 183272
        
        Case 183273
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183274)

    End Select

    Exit Sub

End Sub

Public Sub ClienteBenef_Formata(lCliente As Long)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_ClienteBenef_Formata

    ClienteBenef.Text = lCliente
    
    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(ClienteBenef, objcliente, iCodFilial)
    If lErro <> SUCESSO Then gError 210313

    lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 210314

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", FilialClienteBenef, colCodigoNome)

    
    Exit Sub

Erro_ClienteBenef_Formata:

    Select Case gErr
    
        Case 210313, 210314
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210315)

    End Select

    Exit Sub

End Sub

Public Sub FilialClienteBenef_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialClienteBenef_Formata

    objFilial.Text = CStr(iFilial)
    sCliente = ClienteBenef.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 210315

    If lErro = 17660 Then gError 210316

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_FilialClienteBenef_Formata:

    Select Case gErr

        Case 210315
        
        Case 210316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210317)

    End Select

    Exit Sub

End Sub


Private Function Valida_Cliente() As Long
'Faz a validação do cliente

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFilialCliente As New ClassFilialCliente
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Valida_Cliente

    'Se o campo cliente não foi alterado => sai da função
    If iClienteAlterado = 0 Then Exit Function

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 183308

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 183309

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialCliente, iCodFilial)
                
    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialCliente.Clear
        
    End If
    
    iClienteAlterado = 0
    
    Valida_Cliente = SUCESSO

    Exit Function

Erro_Valida_Cliente:

    Valida_Cliente = gErr
    
    Select Case gErr

        Case 183308, 183309
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183310)

    End Select

    Exit Function

End Function

Private Function Valida_ClienteBenef() As Long
'Faz a validação do cliente

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFilialCliente As New ClassFilialCliente
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Valida_ClienteBenef

    'Se o campo cliente não foi alterado => sai da função
    If iClienteBenefAlterado = 0 Then Exit Function

    'Se Cliente está preenchido
    If Len(Trim(ClienteBenef.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(ClienteBenef, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 210307

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 210308

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialClienteBenef, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialClienteBenef, iCodFilial)
        
    'Se Cliente não está preenchido
    ElseIf Len(Trim(ClienteBenef.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialClienteBenef.Clear
        
    End If
    
    iClienteBenefAlterado = 0
    
    Valida_ClienteBenef = SUCESSO

    Exit Function

Erro_Valida_ClienteBenef:

    Valida_ClienteBenef = gErr
    
    Select Case gErr

        Case 210307, 210308
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210309)

    End Select

    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)

    End If

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If


End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)
    Call Exibe_CampoDet_Grid(objGridItens, objGridItens.objGrid.Col, DetColuna)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataVenda_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataVenda_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataVenda_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataVenda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataVenda
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Servico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Servico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Servico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Servico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Servico
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub UM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Solicitacao_Change()

    iAlterado = REGISTRO_ALTERADO
    DetColuna.Text = Solicitacao.Text
End Sub

Public Sub Solicitacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Solicitacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Solicitacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Solicitacao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Reparo_Change()

    iAlterado = REGISTRO_ALTERADO
    DetColuna.Text = Reparo.Text
End Sub

Public Sub Reparo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Reparo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Reparo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Reparo
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Garantia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Garantia_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Garantia_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Garantia_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Garantia
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Contrato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Contrato_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Contrato_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Contrato_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Contrato
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub StatusItem_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub StatusItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub StatusItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub StatusItem_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = StatusItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataBaixa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataBaixa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataBaixa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataBaixa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataBaixa
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Lote_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Lote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Lote
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub FilialOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub FilialOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = FilialOP
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col
    
            'Se for a de Produto
            Case iGrid_Produto_Col
                lErro = Saida_Celula_Produto(objGridInt)
                If lErro <> SUCESSO Then gError 183333
    
            'Se for a de DataVenda
            Case iGrid_DataVenda_Col
                lErro = Saida_Celula_DataVenda(objGridInt)
                If lErro <> SUCESSO Then gError 183684
    
            'Se for a de Lote
            Case iGrid_Lote_Col
                lErro = Saida_Celula_Lote(objGridInt)
                If lErro <> SUCESSO Then gError 183335
        
            'Se for a de FilialOP
            Case iGrid_FilialOP_Col
                lErro = Saida_Celula_FilialOP(objGridInt)
                If lErro <> SUCESSO Then gError 183336

            'Se for a de Produto
            Case iGrid_Servico_Col
                lErro = Saida_Celula_Servico(objGridInt)
                If lErro <> SUCESSO Then gError 183337
    
            'Se for a de Unidade de Medida
            Case iGrid_UM_Col
                lErro = Saida_Celula_UM(objGridInt)
                If lErro <> SUCESSO Then gError 183338
    
            'Se for a de Quantidade
            Case iGrid_Quantidade_Col
                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 183339
        
            'Se for a de Solicitacao
            Case iGrid_Solicitacao_Col
                lErro = Saida_Celula_Solicitacao(objGridInt)
                If lErro <> SUCESSO Then gError 183340
    
            'Se for a de Solicitacao
            Case iGrid_Reparo_Col
                lErro = Saida_Celula_Padrao(objGridInt, Reparo)
                If lErro <> SUCESSO Then gError 183340
    
            'Se for a de Garantia
            Case iGrid_Garantia_Col
                lErro = Saida_Celula_Garantia(objGridInt)
                If lErro <> SUCESSO Then gError 183341
    
            'Se for a de Contrato
            Case iGrid_Contrato_Col
                lErro = Saida_Celula_Contrato(objGridInt)
                If lErro <> SUCESSO Then gError 183342
    
            Case iGrid_StatusItem_Col
                lErro = Saida_Celula_StatusItem(objGridInt)
                If lErro <> SUCESSO Then gError 195515
    
            'Se for a de DataBaixa
            Case iGrid_DataBaixa_Col
                lErro = Saida_Celula_DataBaixa(objGridInt)
                If lErro <> SUCESSO Then gError 210565
    
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 183343
    
    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 183333 To 183343, 183684, 195515, 210565

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183344)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim lGarantia As Long
Dim sContrato As String
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 183345

        'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
        If lErro <> SUCESSO Then gError 183346
                
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 195706

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
                
    Else
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Solicitacao_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183347

    If Len(Trim(Produto.ClipText)) <> 0 Then

        GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoDesc_Col) = objProduto.sDescricao
    
        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    
        End If
    
        If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then
    
            'se tiver garantia associada, traz para a tela
            lErro = CF("Pesquisa_Garantia", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)), lGarantia)
            If lErro <> SUCESSO Then gError 183524
    
            If lGarantia <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = lGarantia
    
        End If
    
        If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then
    
            'se tiver contrato associada, traz para a tela
            lErro = CF("Pesquisa_Contrato", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)), sContrato)
            If lErro <> SUCESSO Then gError 183525
    
            If Len(Trim(sContrato)) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = sContrato
    
        End If
        
        GridItens.TextMatrix(GridItens.Row, iGrid_StatusItem_Col) = StatusItem.List(0)

    End If

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 183345, 183347, 183524, 183525
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 183346
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 195706
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183348)

    End Select

    Exit Function

End Function

Function Saida_Celula_DataVenda(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Venda que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataVenda

    Set objGridInt.objControle = DataVenda

    If Len(Trim(DataVenda.ClipText)) > 0 Then
        'Critica a Data informada
        lErro = Data_Critica(DataVenda.Text)
        If lErro <> SUCESSO Then gError 183685
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183686

    Saida_Celula_DataVenda = SUCESSO

    Exit Function

Erro_Saida_Celula_DataVenda:

    Saida_Celula_DataVenda = gErr

    Select Case gErr

        Case 183685, 183686
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183687)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Lote(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim dQuantidade As Double
Dim lGarantia As Long
Dim sContrato As String

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote
    
    If Len(Trim(Lote.Text)) > 0 Then
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 183350
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objProduto.sCodigo = sProdutoFormatado
                    
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 183351
                
            If lErro = 28030 Then gError 183352
                
            If gobjSRV.iVerificaLote = VERIFICA_LOTE Then
                
                'Se for rastro por lote
                If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    
                    objRastroLote.sCodigo = Lote.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 183353
                    
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 183354
                    
                'Se for rastro por OP
                ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                    
                    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))) > 0 Then
                        
                        objRastroLote.sCodigo = Lote.Text
                        objRastroLote.sProduto = sProdutoFormatado
                        objRastroLote.iFilialOP = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))
                        
                        'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                        lErro = CF("RastreamentoLote_Le", objRastroLote)
                        If lErro <> SUCESSO And lErro <> 75710 Then gError 183355
                        
                        'Se não encontrou --> Erro
                        If lErro = 75710 Then gError 183356
                        
                    End If
                    
                End If
            
            End If
        
        End If
    
    End If
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183357

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then

        'se tiver garantia associada, traz para a tela
        lErro = CF("Pesquisa_Garantia", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)), lGarantia)
        If lErro <> SUCESSO Then gError 183353

        If lGarantia <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = lGarantia
    
    End If

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then

        'se tiver contrato associada, traz para a tela
        lErro = CF("Pesquisa_Contrato", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)), sContrato)
        If lErro <> SUCESSO Then gError 183519

        If Len(Trim(sContrato)) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = sContrato

    End If

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 183350, 183351, 183353, 183355, 183357, 183519
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 183352
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 183354, 183356
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NUMSERIE_NAO_CADASTRADO", gErr, objRastroLote.sCodigo, objRastroLote.sProduto, objRastroLote.iFilialOP)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183358)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialOP(objGridInt As AdmGrid) As Long
'Faz a saida de celula da Filial da Ordem de Produção

Dim lErro As Long
Dim objFilialOP As New AdmFiliais
Dim iCodigo As Integer
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objRastroLote As New ClassRastreamentoLote
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dQuantidade As Double
Dim lGarantia As Long
Dim sContrato As String

On Error GoTo Erro_Saida_Celula_FilialOP

    Set objGridInt.objControle = FilialOP

    If Len(Trim(FilialOP.Text)) <> 0 Then
            
        'Verifica se é uma FilialOP selecionada
        If FilialOP.Text <> FilialOP.List(FilialOP.ListIndex) Then
        
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOP, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 183359
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOP.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 183360
        
                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 183361
        
                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome
            
            
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 183362
                    
        End If
        
        If gobjSRV.iVerificaLote = VERIFICA_LOTE Then
        
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col))) > 0 Then
                
                lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 183363
                                    
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
                    objRastroLote.sCodigo = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)
                
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 183364
                    
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 183365
                                
                End If
                
            End If
        
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183366

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then

        'se tiver garantia associada, traz para a tela
        lErro = CF("Pesquisa_Garantia", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)), lGarantia)
        If lErro <> SUCESSO Then gError 183520

        If lGarantia <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = lGarantia
    
    End If

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then

        'se tiver contrato associada, traz para a tela
        lErro = CF("Pesquisa_Contrato", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)), sContrato)
        If lErro <> SUCESSO Then gError 183521

        If Len(Trim(sContrato)) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = sContrato

    End If

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 183359, 183360, 183363, 183364, 183366, 183520, 183521
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 183361
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 183362
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 183365
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183367)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Servico(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim lGarantia As Long
Dim sContrato As String
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Servico

    Set objGridInt.objControle = Servico

    If Len(Trim(Servico.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica", Servico.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 183368

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 195707

        Servico.PromptInclude = False
        Servico.Text = sProduto
        Servico.PromptInclude = True

        'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
        If lErro <> SUCESSO Then gError 183369
                
        If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then gError 183383
                
        'Unidade de Medida
        GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = objProduto.sSiglaUMVenda
                
    Else
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183370

    If Len(Trim(Servico.ClipText)) <> 0 Then

        GridItens.TextMatrix(GridItens.Row, iGrid_ServicoDesc_Col) = objProduto.sDescricao
    
        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    
        End If
    
        If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then
    
            'se tiver garantia associada, traz para a tela
            lErro = CF("Pesquisa_Garantia", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)), lGarantia)
            If lErro <> SUCESSO Then gError 183522
    
            If lGarantia <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = lGarantia
        
        End If
    
        If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then
    
            'se tiver contrato associada, traz para a tela
            lErro = CF("Pesquisa_Contrato", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)), sContrato)
            If lErro <> SUCESSO Then gError 183523
    
            If Len(Trim(sContrato)) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = sContrato
    
        End If
    
    End If

    Saida_Celula_Servico = SUCESSO

    Exit Function

Erro_Saida_Celula_Servico:

    Saida_Celula_Servico = gErr

    Select Case gErr

        Case 183368, 183370, 183522, 183523
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 183369
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Servico.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Servico.Text
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 183383
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZA_PROD_NAO_SERVICO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195707
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183371)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidadede Medida que está deixando de ser a corrente

Dim lErro As Long
Dim sUmAnterior As String

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UM

    'Guarda a Unidade de medida anteriormente selecionada
    sUmAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col)

    'Coloca a Um no grid de itens
    GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = UM.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183372

    'Se a Um selecionada agora é diferente da anterior
    If sUmAnterior <> UM.Text And sUmAnterior <> "" Then
    
        lErro = Atualiza_UM(GridItens.Row, sUmAnterior, UM.Text)
        If lErro <> SUCESSO Then gError 183373
    
    End If

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 183372, 183373
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183374)

    End Select

End Function

Private Function Atualiza_UM(ByVal iLinha As Integer, ByVal sUmAnterior As String, ByVal sUMNova As String) As Long
'Atualiza quantidades em funcao de troca de UM

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double, dPrecoUnitario As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Atualiza_UM

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Servico_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 183375

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridItens
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 183376
    
    If lErro = 28030 Then gError 183377

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUmAnterior, sUMNova, dFator)
    If lErro <> SUCESSO Then gError 183378

    'Atualiza o Grid
    GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) * dFator)

    Atualiza_UM = SUCESSO
    
    Exit Function

Erro_Atualiza_UM:

    Atualiza_UM = gErr

    Select Case gErr

        Case 183375, 183376, 183378

        Case 183377
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183379)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 183380

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183381
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 183380, 183381
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183382)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Solicitacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Solicitacao

    Set objGridInt.objControle = Solicitacao

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183440
    
    Saida_Celula_Solicitacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Solicitacao:

    Saida_Celula_Solicitacao = gErr

    Select Case gErr

        Case 183440
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183441)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Garantia(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Garantia está deixando de ser a corrente

Dim lErro As Long
Dim objGarantia As New ClassGarantia

On Error GoTo Erro_Saida_Celula_Garantia

    Set objGridInt.objControle = Garantia

    If Len(Trim(Garantia.Text)) > 0 Then

        lErro = Long_Critica(Garantia.Text)
        If lErro <> SUCESSO Then gError 183570

        objGarantia.iFilialEmpresa = giFilialEmpresa
        objGarantia.lCodigo = StrParaLong(Garantia.Text)
        objGarantia.sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        objGarantia.sServico = GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col)
        objGarantia.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
        objGarantia.iFilialOP = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))

        lErro = CF("Testa_Garantia", objGarantia)
        If lErro <> SUCESSO Then gError 183591
        
    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183573
    
    Saida_Celula_Garantia = SUCESSO

    Exit Function

Erro_Saida_Celula_Garantia:

    Saida_Celula_Garantia = gErr

    Select Case gErr

        Case 183570, 183573, 183591
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183574)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Contrato(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Contrato está deixando de ser a corrente

Dim lErro As Long
Dim objItensDeContratoSrv As New ClassItensDeContratoSrv

On Error GoTo Erro_Saida_Celula_Contrato

    Set objGridInt.objControle = Contrato

    If Len(Trim(Contrato.Text)) > 0 Then

        objItensDeContratoSrv.iFilialEmpresa = giFilialEmpresa
        objItensDeContratoSrv.sCodigoContrato = Contrato.Text
        objItensDeContratoSrv.sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        objItensDeContratoSrv.sServico = GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col)
        objItensDeContratoSrv.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
        objItensDeContratoSrv.iFilialOP = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))
        
        lErro = CF("Testa_Contrato", objItensDeContratoSrv)
        If lErro <> SUCESSO Then gError 183612

    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183613
    
    Saida_Celula_Contrato = SUCESSO

    Exit Function

Erro_Saida_Celula_Contrato:

    Saida_Celula_Contrato = gErr

    Select Case gErr

        Case 183611 To 183613
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183614)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_StatusItem(objGridInt As AdmGrid) As Long
'faz a critica da celula de StatusItem do grid que está deixando de ser a corrente
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_StatusItem

    Set objGridInt.objControle = StatusItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195516

    Saida_Celula_StatusItem = SUCESSO

    Exit Function

Erro_Saida_Celula_StatusItem:

    Saida_Celula_StatusItem = gErr

    Select Case gErr

        Case 195516
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195517)

    End Select

    Exit Function

End Function

Function Saida_Celula_DataBaixa(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Baixa que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataBaixa

    Set objGridInt.objControle = DataBaixa

    If Len(Trim(DataBaixa.ClipText)) > 0 Then
        'Critica a Data informada
        lErro = Data_Critica(DataBaixa.Text)
        If lErro <> SUCESSO Then gError 210566
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 210567

    Saida_Celula_DataBaixa = SUCESSO

    Exit Function

Erro_Saida_Celula_DataBaixa:

    Saida_Celula_DataBaixa = gErr

    Select Case gErr

        Case 210566, 210567
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210568)

    End Select

    Exit Function

End Function


'Private Function Pesquisa_Garantia() As Long
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objGarantia As New ClassGarantia
'Dim sServicoFormatado As String
'Dim iServicoPreenchido As Integer
'
'On Error GoTo Erro_Pesquisa_Garantia
'
'    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then
'
'        'Formata o Produto para o BD
'        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 183470
'
'        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
'
'            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col))) > 0 Then
'
'                objProduto.sCodigo = sProdutoFormatado
'
'                'Lê os demais atributos do Produto
'                lErro = CF("Produto_Le", objProduto)
'                If lErro <> SUCESSO And lErro <> 28030 Then gError 183471
'
'                If lErro = 28030 Then gError 183472
'
'                objGarantia.iFilialEmpresa = giFilialEmpresa
'                objGarantia.sProduto = objProduto.sCodigo
'
'                'Formata o Produto para o BD
'                lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), sServicoFormatado, iServicoPreenchido)
'                If lErro <> SUCESSO Then gError 183733
'
'                If iServicoPreenchido = PRODUTO_PREENCHIDO Then
'
'                    objGarantia.sServico = sServicoFormatado
'
'                    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
'
'                        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col))) > 0 Then
'
'                            objGarantia.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
'
'                            lErro = CF("Garantia_Le_Lote", objGarantia)
'                            If lErro <> SUCESSO And lErro <> 183445 Then gError 183473
'
'                            If lErro = SUCESSO Then
'                                GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = objGarantia.lCodigo
'                            End If
'
'                        End If
'
'                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
'
'                        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col))) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))) > 0 Then
'
'                            objGarantia.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
'                            objGarantia.iFilialOP = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))
'
'                            lErro = CF("Garantia_Le_Lote_FilialOP", objGarantia)
'                            If lErro <> SUCESSO And lErro <> 183460 Then gError 183474
'
'                            If lErro = SUCESSO Then
'                                GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = objGarantia.lCodigo
'                            End If
'
'                        End If
'
'                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
'
'                        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col))) > 0 Then
'
'                            objGarantia.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
'
'                            lErro = CF("Garantia_Le_NumSerie", objGarantia)
'                            If lErro <> SUCESSO And lErro <> 183466 Then gError 183475
'
'                            If lErro = SUCESSO Then
'                                GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = objGarantia.lCodigo
'                            End If
'
'                        End If
'
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
'    End If
'
'    Pesquisa_Garantia = SUCESSO
'
'    Exit Function
'
'Erro_Pesquisa_Garantia:
'
'    Pesquisa_Garantia = gErr
'
'    Select Case gErr
'
'        Case 183470, 183471, 183473 To 183475, 183733
'
'        Case 183472
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183476)
'
'    End Select
'
'    Exit Function
'
'End Function

'Private Function Pesquisa_Contrato() As Long
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objItensContratoSrv As New ClassItensDeContratoSrv
'Dim sServicoFormatado As String
'Dim iServicoPreenchido As Integer
'
'On Error GoTo Erro_Pesquisa_Contrato
'
'    If gobjSRV.iContratoAutoSolic = CONTRATO_AUTOMATICO_SOLICITACAO Then
'
'        'Formata o Produto para o BD
'        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 183512
'
'        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
'
'            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col))) > 0 Then
'
'                objProduto.sCodigo = sProdutoFormatado
'
'                'Lê os demais atributos do Produto
'                lErro = CF("Produto_Le", objProduto)
'                If lErro <> SUCESSO And lErro <> 28030 Then gError 183513
'
'                If lErro = 28030 Then gError 183514
'
'                objItensContratoSrv.iFilialEmpresa = giFilialEmpresa
'                objItensContratoSrv.sProduto = objProduto.sCodigo
'
'                'Formata o Produto para o BD
'                lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), sServicoFormatado, iServicoPreenchido)
'                If lErro <> SUCESSO Then gError 183734
'
'                If iServicoPreenchido = PRODUTO_PREENCHIDO Then
'
'                    objItensContratoSrv.sServico = sServicoFormatado
'
'                    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
'
'                        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col))) > 0 Then
'
'                           objItensContratoSrv.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
'
'                            lErro = CF("ItensDeContratoSrv_Le_Lote", objItensContratoSrv)
'                            If lErro <> SUCESSO And lErro <> 183480 Then gError 183515
'
'                            If lErro = SUCESSO Then
'                                GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = objItensContratoSrv.sCodigoContrato
'                            End If
'
'                        End If
'
'                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
'
'                        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col))) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))) > 0 Then
'
'                            objItensContratoSrv.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
'                            objItensContratoSrv.iFilialOP = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))
'
'                            lErro = CF("ItensDeContratoSrv_Le_Lote_FilialOP", objItensContratoSrv)
'                            If lErro <> SUCESSO And lErro <> 183508 Then gError 183516
'
'                            If lErro = SUCESSO Then
'                                GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = objItensContratoSrv.sCodigoContrato
'                            End If
'
'                        End If
'
'                    ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
'
'                        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col))) > 0 Then
'
'                            objItensContratoSrv.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
'
'                            lErro = CF("ItensDeContratoSrv_Le_NumSerie", objItensContratoSrv)
'                            If lErro <> SUCESSO And lErro <> 183516 Then gError 183517
'
'                            If lErro = SUCESSO Then
'                                GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = objItensContratoSrv.sCodigoContrato
'                            End If
'
'                        End If
'
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
'    End If
'
'    Pesquisa_Contrato = SUCESSO
'
'    Exit Function
'
'Erro_Pesquisa_Contrato:
'
'    Pesquisa_Contrato = gErr
'
'    Select Case gErr
'
'        Case 183512, 183513, 183515 To 183517, 183734
'
'        Case 183514
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183518)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 183616

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOP.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOP.ItemData(FilialOP.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 183616

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183740)

    End Select

    Exit Function

End Function

Private Sub Limpa_SolicSRV()

Dim iIndice As Integer

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Status.Caption = STRING_STATUS_ABERTO
    
    FilialCliente.Clear
    FilialClienteBenef.Clear
    
    'Coloca data atual como padrão
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Limpa a combo de atendentes
    Atendente.ListIndex = -1
    
    'Seleciona o atendente padrão. Atendente padrão é o atendente vinculado ao usuário ativo
    'Para cada atendente da combo AtendenteDe
    For iIndice = 0 To Atendente.ListCount - 1
    
        'Se o conteúdo do atendente for igual ao seu código + "-" + nome reduzido do usuário ativo
        If Atendente.List(iIndice) = Atendente.ItemData(iIndice) & SEPARADOR & gsUsuario Then
        
            'Significa que achou o atendente "default"
            'Seleciona o atendente na combo
            Atendente.ListIndex = iIndice
            
            'Sai do For
            Exit For
        End If
    Next
    
    'Limpa a combo contatos
    CodigoCRM.Caption = ""
    
    Call Grid_Limpa(objGridItens)
    
    Call Define_Padrao
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iClienteBenefAlterado = 0
    iFilialCliBenefAlterada = 0
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    Call Trata_AssuntoCRM
    
    'Verifica se todos os campos obrigatórios estão preenchidos
    lErro = Valida_Gravacao()
    If lErro <> SUCESSO Then gError 183621

    'Move os dados da tela para o objRelacionamentoClie
    lErro = Move_Solicitacao_Memoria(objSolicSRV)
    If lErro <> SUCESSO Then gError 183622

    'Verifica se essa solicitação já existe no BD
    'e, em caso positivo, alerta ao usuário que está sendo feita uma alteração
    lErro = Trata_Alteracao(objSolicSRV, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
    If lErro <> SUCESSO Then gError 183642
    
    'Grava no BD
    lErro = CF("SolicitacaoSRV_Grava", objSolicSRV)
    If lErro <> SUCESSO Then gError 183643

    'Se for para imprimir o relacionamento depois da gravação
    If ImprimeGravacao.Value = vbChecked Then

        'Dispara função para imprimir orçamento
        lErro = SolicSRV_Imprime(objSolicSRV.lCodigo)
        If lErro <> SUCESSO Then gError 183644

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 183621, 183622, 183642, 183643, 183644
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183645)

    End Select

    Exit Function

End Function

Private Function Valida_Gravacao() As Long
'Verifica se os dados da tela são válidos para a gravação do registro

Dim lErro As Long
Dim iIndice As Integer
Dim dQuantidade As Double
Dim sStatusItem As String
Dim bAchou As Boolean
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Valida_Gravacao

    bAchou = False
    
    'Se o código não estiver preenchido => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 183616
    
    'Se a data não estiver preenchida => erro
    If Len(Trim(Data.ClipText)) = 0 Then gError 183617
    
    'Se o cliente não estiver preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 183618
    
    'Se a filial do cliente não estiver preenchida => erro
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 183619
    
    If objGridItens.iLinhasExistentes = 0 Then gError 183641
    
    For iIndice = 1 To objGridItens.iLinhasExistentes

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Then gError 183688
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DataVenda_Col))) = 0 Then gError 183689
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Servico_Col))) = 0 Then gError 183690
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UM_Col))) = 0 Then gError 183691
        
        dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        If dQuantidade = 0 Then gError 183692
        
        sStatusItem = GridItens.TextMatrix(iIndice, iGrid_StatusItem_Col)
        
        'Verifica se é uma ordem de servico baixada
        If Status.Caption = STRING_STATUS_BAIXADO Then
            
            'Verifica se o Status é Normal para o item
            If UCase(sStatusItem) <> UCase(STRING_BAIXADA) Then bAchou = True
                
        End If
        
    
    Next
    
    'se a OS está baixada e existe item com situacao='normal'
    If bAchou = True Then
    
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_REATIVACAO_SOLICITACAOSRV", Codigo.Text)
        'se não for reativar a OS sai da gravação
        If vbMsg = vbNo Then gError 195513
    
    ElseIf bAchou = False And Status.Caption = STRING_STATUS_BAIXADO Then
        gError 195514
    End If
        
    Valida_Gravacao = SUCESSO

    Exit Function

Erro_Valida_Gravacao:

    Valida_Gravacao = gErr
    
    Select Case gErr
    
        Case 183616
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 183617
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            
        Case 183618
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 183619
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 183641
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_PREENCHIDA", gErr)
        
        Case 183688
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, iIndice)
        
        Case 183689
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENDA_NAO_PREENCHIDA_GRID", gErr, iIndice)
        
        Case 183690
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case 183691
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 183692
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA_GRID1", gErr, iIndice)

        Case 195513

        Case 195514
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRVBAIXADA_NAO_REATIVADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183620)

    End Select

End Function

Private Function Move_Solicitacao_Memoria(objSolicSRV As ClassSolicSRV) As Long
'Move os dados da tela para objSolicSRV

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim objclienteBenef As New ClassCliente
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Move_Solicitacao_Memoria

    objSolicSRV.iFilialEmpresa = giFilialEmpresa
    
    objSolicSRV.lCodigo = StrParaLong(Codigo.Text)

    objSolicSRV.dtData = StrParaDate(Data.Text)
    
    If Len(Trim(Hora.ClipText)) > 0 Then
        objSolicSRV.dtHora = CDate(Hora.Text)
    Else
        objSolicSRV.dtHora = Time
    End If
    
    objcliente.sNomeReduzido = Cliente.Text

    'Lê o Cliente através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 183623

    'Se não achou o Cliente --> erro
    If lErro = 12348 Then gError 183624

    'Guarda código do Cliente em objPedidoVenda
    objSolicSRV.lCliente = objcliente.lCodigo

    objSolicSRV.iFilial = Codigo_Extrai(FilialCliente.Text)
    
    If Len(ClienteBenef.Text) > 0 Then
    
        objclienteBenef.sNomeReduzido = ClienteBenef.Text
    
        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objclienteBenef)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 210319
    
        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 210320

        'Guarda código do Cliente em objPedidoVenda
        objSolicSRV.lClienteBenef = objclienteBenef.lCodigo
    
        objSolicSRV.iFilialClienteBenef = Codigo_Extrai(FilialClienteBenef.Text)
    
    End If
    
    
    objSolicSRV.iAtendente = Codigo_Extrai(Atendente.Text)
    
    'Verifica se vendedor existe
    If Len(Trim(Vendedor.Text)) > 0 Then
        
        objVendedor.sNomeReduzido = Trim(Vendedor.Text)

        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 183625

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 183626

        objSolicSRV.iVendedor = objVendedor.iCodigo

    End If
    
    objSolicSRV.sOBS = Obs.Text
    objSolicSRV.iPrazo = StrParaInt(Prazo.Text)
    If OptPrazoUteis.Value Then
        objSolicSRV.iPrazoTipo = 0
    Else
        objSolicSRV.iPrazoTipo = 1
    End If
    objSolicSRV.dtDataEntrega = StrParaDate(DataEntrega.Text)

    'Move Grid Itens para memória
    lErro = Move_GridItens_Memoria(objSolicSRV)
    If lErro <> SUCESSO Then gError 183627
    
    If GravarCRM.Value = vbChecked Then
        objSolicSRV.iGravarCRM = MARCADO
    
        lErro = Move_CRM_Memoria(objSolicSRV)
        If lErro <> SUCESSO Then gError 183627
    
    Else
        objSolicSRV.iGravarCRM = DESMARCADO
    End If

    If Tipo.ListIndex <> -1 Then
        objSolicSRV.lTipo = Tipo.ItemData(Tipo.ListIndex)
    End If
    
        If Fase.ListIndex <> -1 Then
        objSolicSRV.lFase = Fase.ItemData(Fase.ListIndex)
    End If
    
    Move_Solicitacao_Memoria = SUCESSO

    Exit Function

Erro_Move_Solicitacao_Memoria:

    Move_Solicitacao_Memoria = gErr

    Select Case gErr

        Case 183623, 183625, 183627, 210319

        Case 183624
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 183626
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)

        Case 210320
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEBENEF_NAO_CADASTRADO", gErr, ClienteBenef.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183628)

    End Select

    Exit Function

End Function

Private Function Move_GridItens_Memoria(objSolicSRV As ClassSolicSRV) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim sProduto As String
Dim sServico As String
Dim iPreenchido As Integer
Dim dQuantidade As Double
Dim objItensSolicSRV As ClassItensSolicSRV
Dim iIndice As Integer
Dim sStatusItem As String
Dim iCount As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItensSolicSRV = New ClassItensSolicSRV
    
        'Formata o produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 183629
    
        If iPreenchido = PRODUTO_VAZIO Then gError 183630
    
        objItensSolicSRV.sProduto = sProduto
    
        'Formata o serviço
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Servico_Col), sServico, iPreenchido)
        If lErro <> SUCESSO Then gError 183631
    
        If iPreenchido = PRODUTO_VAZIO Then gError 183632
    
        objItensSolicSRV.sServico = sServico
    
        'Armazena os dados do item
        objItensSolicSRV.sUM = GridItens.TextMatrix(iIndice, iGrid_UM_Col)
        objItensSolicSRV.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        lErro = Move_RastroEstoque_Memoria(iIndice, objItensSolicSRV)
        If lErro <> SUCESSO Then gError 183639
        
        objItensSolicSRV.sSolicitacao = GridItens.TextMatrix(iIndice, iGrid_Solicitacao_Col)
        objItensSolicSRV.sReparo = GridItens.TextMatrix(iIndice, iGrid_Reparo_Col)
        objItensSolicSRV.lGarantia = StrParaLong(GridItens.TextMatrix(iIndice, iGrid_Garantia_Col))
        objItensSolicSRV.sContrato = GridItens.TextMatrix(iIndice, iGrid_Contrato_Col)
        objItensSolicSRV.dtDataVenda = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataVenda_Col))
        objItensSolicSRV.lNumIntDoc = StrParaLong(GridItens.TextMatrix(iIndice, iGrid_NumIntDoc_Col))
    
        'Seleciona o status
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_StatusItem_Col))) > 0 Then
            sStatusItem = GridItens.TextMatrix(iIndice, iGrid_StatusItem_Col)
            For iCount = 0 To StatusItem.ListCount - 1
                If StatusItem.List(iCount) = sStatusItem Then
                    objItensSolicSRV.iStatusItem = StatusItem.ItemData(iCount)
                    Exit For
                End If
            Next
        End If
    
        objItensSolicSRV.dtDataBaixa = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataBaixa_Col))
    
    
        objSolicSRV.colItens.Add objItensSolicSRV
    
    Next
    
    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 183629, 183631, 183639

        Case 183630
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case 183632
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183640)

    End Select

    Exit Function

End Function

Private Function Move_RastroEstoque_Memoria(iLinha As Integer, objItensSolicSRV As ClassItensSolicSRV) As Long
'Move o Rastro dos Itens de Movimento

Dim objProduto As New ClassProduto, lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_RastroEstoque_Memoria
    
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 183633
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 183634
    
        If lErro = 28030 Then gError 183635
        
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            
            'Se colocou o Número do Lote
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
                objItensSolicSRV.sLote = GridItens.TextMatrix(iLinha, iGrid_Lote_Col)
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            'se o lote está preenchido e a filial não ==> erro
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
               
                If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_FilialOP_Col))) = 0 Then gError 183636
                
                objItensSolicSRV.sLote = GridItens.TextMatrix(iLinha, iGrid_Lote_Col)
                objItensSolicSRV.iFilialOP = Codigo_Extrai(GridItens.TextMatrix(iLinha, iGrid_FilialOP_Col))
                
            End If
                
            'se a filial está preenchida e o lote não ==> erro
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_FilialOP_Col))) <> 0 And _
               Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Lote_Col))) = 0 Then gError 183637
                
        End If
    
    End If
    
    Move_RastroEstoque_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_RastroEstoque_Memoria:

    Move_RastroEstoque_Memoria = gErr
    
    Select Case gErr
        
        Case 183633, 183634
        
        Case 183635
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 183636
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
        
        Case 183637
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183638)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer
Dim sUnidadeMed As String
Dim iIndice As Integer
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 183678

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Servico_Col), sServicoFormatado, iServicoPreenchido)
    If lErro <> SUCESSO Then gError 183679

    Select Case objControl.Name
        
        Case Produto.Name
            'Se o numintdoc estiver preenchido desabilita
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_NumIntDoc_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If



        Case DataVenda.Name, Lote.Name, FilialOP.Name, Servico.Name, Garantia.Name, Contrato.Name
            If iProdutoPreenchido = PRODUTO_VAZIO Or Len(Trim(GridItens.TextMatrix(iLinha, iGrid_NumIntDoc_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        
        Case UM.Name
            'guarda a um go grid nessa coluna
            sUM = GridItens.TextMatrix(iLinha, iGrid_UM_Col)
            
            'Se o numintdoc estiver preenchido desabilita
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_NumIntDoc_Col))) > 0 Then
                UM.Enabled = False
            Else
                UM.Enabled = True
            End If
            
            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UM.Text
            
            UM.Clear

            If iServicoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sServicoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 183680

                If lErro = 28030 Then gError 183681

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 183682
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    UM.AddItem objUM.sSigla
                Next
                
                'Tento selecionar na Combo a Unidade anterior
                If UM.ListCount <> 0 Then
    
                    For iIndice = 0 To UM.ListCount - 1
    
                        If UM.List(iIndice) = sUnidadeMed Then
                            UM.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If
            
            Else
                UM.Enabled = False
            End If

        Case Quantidade.Name
            'Se o produto estiver preenchido, habilita o controle
            If iServicoPreenchido = PRODUTO_VAZIO Or Len(Trim(GridItens.TextMatrix(iLinha, iGrid_NumIntDoc_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Solicitacao.Name, StatusItem.Name, Reparo.Name, DataBaixa.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 183678, 183679, 183680, 183682

        Case 183681
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183683)

    End Select

    Exit Sub

End Sub

'**** TRATAMENTO DO SISTEMA DE SETAS - INÍCIO ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objSolicSRV As New ClassSolicSRV
Dim objCampoValor As AdmCampoValor
Dim lErro As Long
Dim lCliente As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "SolicitacaoSRV"

    'Guarda no obj os dados que serão usados para identifica o registro a ser exibido
    objSolicSRV.lCodigo = StrParaDbl(Trim(Codigo.Text))
    objSolicSRV.iFilialEmpresa = giFilialEmpresa
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objSolicSRV.lCodigo, 0, "Codigo"
    colCampoValor.Add "FilialEmpresa", objSolicSRV.iFilialEmpresa, 0, "FilialEmpresa"
    
    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183741)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_Tela_Preenche

    'Guarda o código do campo em questão no obj
    objSolicSRV.lCodigo = colCampoValor.Item("Codigo").vValor
    objSolicSRV.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor

    lErro = Traz_SolicSRV_Tela(objSolicSRV)
    If lErro <> SUCESSO Then gError 183742

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 183742
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183743)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub
'**** FIM DO TRATAMENTO DO SISTEMA DE SETAS ****


Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se o código da Solicitacao não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 183735

    'Dispara função para imprimir relacionamento
    lErro = SolicSRV_Imprime(StrParaLong(Codigo.Text))
    If lErro <> SUCESSO Then gError 183736

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 183735
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 183736

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183737)

    End Select

    Exit Sub

End Sub

Private Function SolicSRV_Imprime(ByVal lCodigo As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_SolicSRV_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Guarda no obj o código da solicitacao passado como parâmetro
    objSolicSRV.lCodigo = lCodigo
    
    'Guarda a FilialEmpresa ativa como filial do relacionamento
    objSolicSRV.iFilialEmpresa = giFilialEmpresa
    
    'Lê os dados da solicitacao para verificar se o mesmo existe no BD
    lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
    If lErro <> SUCESSO And lErro <> 183258 Then gError 183738

    'Se não encontrou => erro, pois não é possível imprimir uma solicitacao inexistente
    If lErro <> SUCESSO Then gError 183739
    
    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDireto("Solicitação de Serviço", "Codigo = @NSSDE", 1, , "NSSDE", CStr(lCodigo), "NSSATE", CStr(lCodigo), "TPRODINIC", "", "TPRODFIM", "", "TCLIENTEINIC", "", "TCLIENTEFIM", "", "NSTATUS", 0, "TATENDDE", "", "TATENDATE", "", "DINIC", CStr(DATA_NULA), "DFIM", CStr(DATA_NULA), "TVENDINIC", "", "TVENDFIM", "")
    If lErro <> SUCESSO Then gError 183740

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    SolicSRV_Imprime = SUCESSO
    
    Exit Function

Erro_SolicSRV_Imprime:

    SolicSRV_Imprime = gErr
    
    Select Case gErr
    
        Case 183738, 183740
        
        Case 183739
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183741)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function

Private Sub BotaoGarantia_Click()

Dim lErro As Long
Dim objGarantia As New ClassGarantia
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoGarantia_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 195714

    If Me.ActiveControl Is Garantia Then
        objGarantia.lCodigo = StrParaLong(Garantia.Text)
    Else
        objGarantia.lCodigo = StrParaLong(GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col))
    End If

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195715

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 195716
    
    colSelecao.Add sProdutoFormatado
    
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195715

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 195716
    
    colSelecao.Add sProdutoFormatado
    colSelecao.Add MARCADO
    
    Call Chama_Tela("GarantiaLista", colSelecao, objGarantia, objEventoGarantia, "Produto = ? AND (NumIntDoc IN (SELECT NumIntGarantia FROM GarantiaProduto WHERE Produto = ?) OR GarantiaTotal = ?)")

    Exit Sub

Erro_BotaoGarantia_Click:

    Select Case gErr

        Case 195714
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 195716
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, GridItens.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195717)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoContrato_Click()

Dim lErro As Long
Dim objContrato As New ClassContrato
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoContrato_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 195714

    If Me.ActiveControl Is Contrato Then
        objContrato.sCodigo = Contrato.Text
    Else
        objContrato.sCodigo = GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col)
    End If

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195715

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 195716
    
    colSelecao.Add sProdutoFormatado
    
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 195715

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 195716
    
    colSelecao.Add MARCADO
    colSelecao.Add sProdutoFormatado
    
    Call Chama_Tela("ContratosLista", colSelecao, objContrato, objEventoContrato, "NumIntDoc IN (SELECT IC.NumIntContrato FROM ItensDeContrato AS IC, ItensDeContratoSRV AS ICS WHERE IC.NumIntDoc = ICS.NumIntItemContrato AND IC.Produto = ? AND (ICS.GarantiaTotal = ? OR ICS.NumIntDoc IN (SELECT NumIntItemContratoSRV FROM ItensDeContratoSRVProd WHERE Produto = ?)) )")

    Exit Sub

Erro_BotaoContrato_Click:

    Select Case gErr

        Case 195714
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 195716
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, GridItens.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195717)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoGarantia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objGarantia As ClassGarantia
Dim iCodigo As Integer

On Error GoTo Erro_objEventoGarantia_evSelecao

    Set objGarantia = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    Garantia.PromptInclude = False
    Garantia.Text = objGarantia.lCodigo
    Garantia.PromptInclude = True

    If Not (Me.ActiveControl Is Garantia) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = Garantia.Text
    
        objGarantia.iFilialEmpresa = giFilialEmpresa
        objGarantia.sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        objGarantia.sServico = GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col)
        objGarantia.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
        objGarantia.iFilialOP = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))

        lErro = CF("Testa_Garantia", objGarantia)
        If lErro <> SUCESSO Then gError 183591
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoGarantia_evSelecao:

    GridItens.TextMatrix(GridItens.Row, iGrid_Garantia_Col) = ""

    Select Case gErr

        Case 183591

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195718)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContrato_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objContrato As ClassContrato
Dim iCodigo As Integer
Dim objItensDeContratoSrv As New ClassItensDeContratoSrv

On Error GoTo Erro_objEventoContrato_evSelecao

    Set objContrato = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    Contrato.Text = objContrato.sCodigo

    If Not (Me.ActiveControl Is Contrato) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = Contrato.Text
    
        objItensDeContratoSrv.iFilialEmpresa = giFilialEmpresa
        objItensDeContratoSrv.sCodigoContrato = Contrato.Text
        objItensDeContratoSrv.sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        objItensDeContratoSrv.sServico = GridItens.TextMatrix(GridItens.Row, iGrid_Servico_Col)
        objItensDeContratoSrv.sLote = GridItens.TextMatrix(GridItens.Row, iGrid_Lote_Col)
        objItensDeContratoSrv.iFilialOP = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))
        
        lErro = CF("Testa_Contrato", objItensDeContratoSrv)
        If lErro <> SUCESSO Then gError 183591
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoContrato_evSelecao:

    GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = ""

    Select Case gErr

        Case 183591

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195718)

    End Select

    Exit Sub

End Sub

Private Sub BotaoOrcamento_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_BotaoOrcamento_Click

    'Se o código não foi preenchido => erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 183715

    objSolicSRV.lCodigo = StrParaLong(Codigo.Text)
    objSolicSRV.iFilialEmpresa = giFilialEmpresa
    
    'Lê os dados da solicitacao para verificar se o mesmo existe no BD
    lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
    If lErro <> SUCESSO And lErro <> 183258 Then gError 183738

    'Se não encontrou => erro, pois não é possível imprimir uma solicitacao inexistente
    If lErro <> SUCESSO Then gError 183739

    colSelecao.Add objSolicSRV.lNumIntDoc

    Call Chama_Tela("OrcamentoSRV1Lista", colSelecao, Nothing, Nothing, "NumIntSolicSRV = ?")

    Exit Sub

Erro_BotaoOrcamento_Click:

    Select Case gErr

        Case 183715
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 183738
        
        Case 183739
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186264)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoOS_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_BotaoOS_Click

    'Se o código não foi preenchido => erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 183715

    objSolicSRV.lCodigo = StrParaLong(Codigo.Text)
    objSolicSRV.iFilialEmpresa = giFilialEmpresa
    
    'Lê os dados da solicitacao para verificar se o mesmo existe no BD
    lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
    If lErro <> SUCESSO And lErro <> 183258 Then gError 183738

    'Se não encontrou => erro, pois não é possível imprimir uma solicitacao inexistente
    If lErro <> SUCESSO Then gError 183739

    colSelecao.Add objSolicSRV.lCodigo

    Call Chama_Tela("OSLista", colSelecao, Nothing, Nothing, "CodSolSrv = ?")

    Exit Sub

Erro_BotaoOS_Click:

    Select Case gErr

        Case 183715
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 183738
        
        Case 183739
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186264)

    End Select

    Exit Sub
    
End Sub

Private Function Carrega_CampoGenerico(ByVal objComboBox As ComboBox, ByVal lCodigo As Long, iPadrao As Integer) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_CampoGenerico

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", lCodigo, objComboBox)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    iPadrao = objComboBox.ListIndex

    Carrega_CampoGenerico = SUCESSO

    Exit Function

Erro_Carrega_CampoGenerico:

    Carrega_CampoGenerico = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Function

End Function

Private Sub Define_Padrao()

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Define_Padrao

    lErro = CF("Config_Le", "SRVConfig", "SOLSRV_TIPOPRAZO_PADRAO", EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sConteudo) = 0 Then
        OptPrazoUteis.Value = True
    Else
        OptPrazoCorr.Value = False
    End If

    lErro = CF("Config_Le", "SRVConfig", "SOLSRV_GRAVA_CRM", EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sConteudo) = MARCADO Then
        GravarCRM.Value = vbChecked
    Else
        GravarCRM.Value = vbUnchecked
    End If
    Call Trata_GravarCRM
    
    lErro = CF("Config_Le", "SRVConfig", "SOLSRV_PRAZO_PADRAO", EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Prazo.PromptInclude = False
    If StrParaInt(sConteudo) > 0 Then
        Prazo.Text = sConteudo
    Else
        Prazo.Text = ""
    End If
    Prazo.PromptInclude = True
    
    Call Carrega_CampoGenerico(StatusCRM, CAMPOSGENERICOS_STATUSRELACCLI, iStatus_ListIndex_Padrao)
    Call Carrega_CampoGenerico(MotivoCRM, CAMPOSGENERICOS_RELACCLI_MOTIVO, iMotivo_ListIndex_Padrao)
    Call Carrega_CampoGenerico(SatisfacaoCRM, CAMPOSGENERICOS_RELACCLI_SATIS, iSatisfacao_ListIndex_Padrao)

    Call Trata_Prazo
    
    MsgAutoCRM.Value = vbChecked
    Encerrado.Value = vbUnchecked
    Call Trata_MsgAutoCRM
    
    Fase.ListIndex = glFasePadrao
    Tipo.ListIndex = glTipoPadrao

    Exit Sub

Erro_Define_Padrao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Sub

End Sub

Private Sub Trata_Prazo()
'Calcula a data de entrega de acordo com o prazo e fator
Dim lErro As Long
Dim objCondicaoPagto  As New ClassCondicaoPagto
Dim objParc As New ClassCondicaoPagtoParc

On Error GoTo Erro_Trata_Prazo

    objCondicaoPagto.dtDataRef = StrParaDate(Data.Text)
    objCondicaoPagto.colParcelas.Add objParc
    
    objParc.iDias = StrParaInt(Prazo.Text)
    objParc.iTipoIntervalo = IIf(OptPrazoCorr.Value, CONDPAGTO_TIPOINTERVALO_DIAS, CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS)
    objParc.dtVencimento = DATA_NULA

    If objCondicaoPagto.dtDataRef <> DATA_NULA And objParc.iTipoIntervalo > 0 Then
        'Só calcula as datas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True, False)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Call DateParaMasked(DataEntrega, objParc.dtVencimento)

    Exit Sub

Erro_Trata_Prazo:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Sub
    
End Sub

Public Sub DataEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrega_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEntrega, iAlterado)
End Sub

Public Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrega_Validate

    'Verifica se a Data foi digitada
    If Len(Trim(DataEntrega.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEntrega.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Trata_DataEntrega

    Exit Sub

Erro_DataEntrega_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183280)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnt_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEnt_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEntrega, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Trata_DataEntrega

    Exit Sub

Erro_UpDownDataEnt_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183737)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnt_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnt_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntrega, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Trata_DataEntrega

    Exit Sub

Erro_UpDownDataEnt_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183739)

    End Select

    Exit Sub

End Sub

Private Sub Trata_DataEntrega()

Dim lErro As Long

On Error GoTo Erro_Trata_DataEntrega

    If StrParaDate(DataEntrega.Text) <> gdtDataEntregaAnt And StrParaDate(DataEntrega.Text) <> DATA_NULA Then
    
        gdtDataEntregaAnt = StrParaDate(DataEntrega.Text)
        
        OptPrazoCorr.Value = True
        
        Prazo.PromptInclude = False
        Prazo.Text = DateDiff("d", StrParaDate(DataEntrega.Text), StrParaDate(Data.Text))
        Prazo.PromptInclude = True
    
    End If

    Exit Sub

Erro_Trata_DataEntrega:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Sub
    
End Sub

Private Sub Trata_AssuntoCRM()

Dim lErro As Long
Dim iIndice As Integer, sTexto As String

On Error GoTo Erro_Trata_AssuntoCRM

    If MsgAutoCRM.Value = vbChecked Then
    
        If Len(Trim(Codigo.Text)) > 0 Then sTexto = "Código SS: " & Codigo.Text
    
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            sTexto = sTexto & vbNewLine & vbNewLine
        
            sTexto = sTexto & "Produto: " & Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col)) & SEPARADOR & GridItens.TextMatrix(iIndice, iGrid_ProdutoDesc_Col)
            sTexto = sTexto & " | Série: " & GridItens.TextMatrix(iIndice, iGrid_Lote_Col)
            sTexto = sTexto & vbNewLine & "Solicitação: " & GridItens.TextMatrix(iIndice, iGrid_Solicitacao_Col)
    
        Next
        
        AssuntoCRM.Text = sTexto
    
    End If

    Exit Sub

Erro_Trata_AssuntoCRM:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Sub
    
End Sub

Private Function Traz_CRM_Tela(ByVal objSolicSRV As ClassSolicSRV) As Long

Dim lErro As Long, iIndice As Integer
Dim objRelacCli As New ClassRelacClientes

On Error GoTo Erro_Traz_CRM_Tela

    CodigoCRM.Caption = ""

    objRelacCli.iTipoDoc = RELACCLI_TIPODOC_SOLSRV
    objRelacCli.lNumIntDocOrigem = objSolicSRV.lNumIntDoc
    
    lErro = CF("RelacCli_Le_TipoDoc", objRelacCli)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
  
    If lErro = SUCESSO Then
    
        CodigoCRM.Caption = CStr(objRelacCli.lCodigo)
    
        If objRelacCli.iStatusCG <> 0 Then Call Combo_Seleciona_ItemData(StatusCRM, objRelacCli.iStatusCG)
        If objRelacCli.lMotivo <> 0 Then Call Combo_Seleciona_ItemData(MotivoCRM, objRelacCli.lMotivo)
        If objRelacCli.lSatisfacao <> 0 Then Call Combo_Seleciona_ItemData(SatisfacaoCRM, objRelacCli.lSatisfacao)
        
        AssuntoCRM.Text = objRelacCli.sAssunto1 & objRelacCli.sAssunto2
    
        'Status
         If objRelacCli.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO Then Encerrado.Value = vbChecked
         
         'DataFim
         'se a DataFim foi preenchida
         If objRelacCli.dtDataFim <> DATA_NULA Then
             DataFim.PromptInclude = False
             DataFim.Text = Format(objRelacCli.dtDataFim, "dd/mm/yy")
             DataFim.PromptInclude = True
         End If
         
         'HoraFim
         'Se a HoraFim foi gravada no BD
         If objRelacCli.dtHoraFim <> 0 Then
             HoraFim.PromptInclude = False
             HoraFim.Text = Format(objRelacCli.dtHoraFim, "hh:mm:ss")
             HoraFim.PromptInclude = True
         End If
    
    End If
    
    Traz_CRM_Tela = SUCESSO

    Exit Function

Erro_Traz_CRM_Tela:

    Traz_CRM_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183268)

    End Select

    Exit Function

End Function

Private Sub DataFim_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataFim, iAlterado)
End Sub

Private Sub DataFim_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lCliente As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_DataFim_Validate

    'Se a data não foi preenchida => sai da função
    If Len(Trim(DataFim.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataFim.Text)
    If lErro <> SUCESSO Then gError 102510
   
    Exit Sub
    
Erro_DataFim_Validate:

    Cancel = True

    Select Case gErr
    
        Case 102510
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166600)
        
    End Select

End Sub

Public Sub HoraFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraFim_Validate

    'Verifica se a hora de saida foi digitada
    If Len(Trim(HoraFim.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraFim.Text)
    If lErro <> SUCESSO Then gError 102511

    Exit Sub

Erro_HoraFim_Validate:

    Cancel = True

    Select Case gErr

        Case 102511

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166601)

    End Select

    Exit Sub

End Sub

Private Sub DataFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub HoraFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Encerrado_Click()
    iAlterado = REGISTRO_ALTERADO
    If Encerrado.Value = vbChecked Then
        FrameFim.Enabled = True
        DataFim.PromptInclude = False
        DataFim.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataFim.PromptInclude = True
        HoraFim.PromptInclude = False
        HoraFim.Text = Format(Time, "hh:mm:ss")
        HoraFim.PromptInclude = True
    Else
        FrameFim.Enabled = False
        DataFim.PromptInclude = False
        DataFim.Text = ""
        DataFim.PromptInclude = True
        HoraFim.PromptInclude = False
        HoraFim.Text = ""
        HoraFim.PromptInclude = True
    End If
End Sub

Private Function Move_CRM_Memoria(ByVal objSolicSRV As ClassSolicSRV) As Long

Dim lErro As Long, iIndice As Integer
Dim objRelacCli As New ClassRelacClientes
Dim vbResult As VbMsgBoxResult
Dim objSolSrvBD As New ClassSolicSRV

On Error GoTo Erro_Move_CRM_Memoria
    
    objRelacCli.iFilialEmpresa = giFilialEmpresa
    objRelacCli.lCodigo = StrParaLong(CodigoCRM.Caption)
    
    lErro = CF("RelacionamentoClientes_Le", objRelacCli)
    If lErro <> SUCESSO And lErro <> 102508 Then gError ERRO_SEM_MENSAGEM
  
    If lErro = SUCESSO Then
    
        If objRelacCli.iTipoDoc <> RELACCLI_TIPODOC_SOLSRV Or objRelacCli.lNumIntDocOrigem = 0 Then gError 211221

        objSolSrvBD.lNumIntDoc = objRelacCli.lNumIntDocOrigem
        
        lErro = CF("SolicitacaoSRV_Le_NumIntDoc", objSolSrvBD)
        If lErro <> SUCESSO And lErro <> 186988 Then gError ERRO_SEM_MENSAGEM
        
        If objSolSrvBD.lCodigo <> objSolicSRV.lCodigo Or lErro <> SUCESSO Then gError 211222
    
    Else
        objRelacCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE
        objRelacCli.lTipo = TIPO_RELACIONAMENTO_SOLSRV
        objRelacCli.dtDataPrevReceb = DATA_NULA
        objRelacCli.dtDataProxCobr = DATA_NULA
    End If
    
    If objSolicSRV.iAtendente = 0 Then gError 211315
        
    objRelacCli.dtData = objSolicSRV.dtData
    objRelacCli.dtHora = objSolicSRV.dtHora
    objRelacCli.lCliente = objSolicSRV.lCliente
    objRelacCli.iFilialCliente = objSolicSRV.iFilial
    objRelacCli.iAtendente = objSolicSRV.iAtendente
    If Len(AssuntoCRM.Text) > 255 Then
        objRelacCli.sAssunto1 = left(AssuntoCRM.Text, 255)
        objRelacCli.sAssunto2 = Mid(AssuntoCRM.Text, 256)
    Else
        objRelacCli.sAssunto1 = AssuntoCRM.Text
        objRelacCli.sAssunto2 = ""
    End If
    objRelacCli.iStatus = Encerrado.Value
    If StatusCRM.ListIndex <> -1 Then objRelacCli.iStatusCG = StatusCRM.ItemData(StatusCRM.ListIndex)
    If MotivoCRM.ListIndex <> -1 Then objRelacCli.lMotivo = MotivoCRM.ItemData(MotivoCRM.ListIndex)
    If SatisfacaoCRM.ListIndex <> -1 Then objRelacCli.lSatisfacao = SatisfacaoCRM.ItemData(SatisfacaoCRM.ListIndex)
    objRelacCli.dtDataFim = StrParaDate(DataFim.Text)
    If Len(Trim(HoraFim.ClipText)) > 0 Then
        objRelacCli.dtHoraFim = StrParaDate(HoraFim.Text)
    End If
    
    Set objSolicSRV.objRelacCli = objRelacCli
    
    Move_CRM_Memoria = SUCESSO

    Exit Function

Erro_Move_CRM_Memoria:

    Move_CRM_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 211221, 211222
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERV_NAO_VINC_RELACCLI", gErr, CodigoCRM.Caption, objSolicSRV.lCodigo)

        Case 211315
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183268)

    End Select

    Exit Function

End Function

Private Sub BotaoCRM_Click()

Dim lErro As Long
Dim objRelacCli As New ClassRelacClientes
    
On Error GoTo Erro_BotaoCRM_Click

    If StrParaLong(CodigoCRM.Caption) <> 0 Then
    
        objRelacCli.lCodigo = StrParaLong(CodigoCRM.Caption)
        objRelacCli.iFilialEmpresa = giFilialEmpresa
    
        Call Chama_Tela("RelacionamentoClientes", objRelacCli)
        
    End If

    Exit Sub

Erro_BotaoCRM_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166625)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Exibe_CampoDet_Grid(ByVal objGridInt As AdmGrid, ByVal iColunaExibir As Integer, ByVal objControle As Object)

On Error GoTo Erro_Exibe_CampoDet_Grid

    giLinhaDet = objGridInt.objGrid.Row
    giColunaDet = iColunaExibir
    objControle.Locked = True
    LinhaDet.Caption = ""
    ColunaDet.Caption = ""
    objControle.Text = ""
    
    If giLinhaDet > 0 And giLinhaDet <= objGridInt.iLinhasExistentes Then
        If giColunaDet = iGrid_Solicitacao_Col Or giColunaDet = iGrid_Reparo_Col Then
            LinhaDet.Caption = CStr(giLinhaDet)
            ColunaDet.Caption = objGridInt.objGrid.TextMatrix(0, giColunaDet)
            objControle.Text = objGridInt.objGrid.TextMatrix(giLinhaDet, giColunaDet)
            objControle.Locked = False
        End If
    End If

    Exit Sub

Erro_Exibe_CampoDet_Grid:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208641)

    End Select

    Exit Sub
    
End Sub

Private Sub DetColuna_Validate(Cancel As Boolean)
    If giLinhaDet <> 0 And giColunaDet <> 0 And Not DetColuna.Locked Then objGridItens.objGrid.TextMatrix(giLinhaDet, giColunaDet) = DetColuna.Text
End Sub

Private Sub Trata_MsgAutoCRM()

    If MsgAutoCRM.Value = vbChecked Then
        AssuntoCRM.Locked = True
        Call Trata_AssuntoCRM
    Else
        AssuntoCRM.Locked = False
    End If

End Sub

Private Sub Trata_GravarCRM()

    If GravarCRM.Value = vbChecked Then
        FrameCRM1.Enabled = True
        FrameCRM2.Enabled = True
    Else
        FrameCRM1.Enabled = False
        FrameCRM2.Enabled = False
        AssuntoCRM.Locked = False
    End If

End Sub

Public Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Tipo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Fase_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Fase_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

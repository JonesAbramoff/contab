VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ProjetoOcx 
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   2
      Left            =   90
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   9225
      Begin VB.CommandButton BotaoVerKit 
         Caption         =   "Ver Kit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   6675
         TabIndex        =   23
         ToolTipText     =   "Visualiza o Kit do Produto"
         Top             =   4080
         Width           =   1155
      End
      Begin VB.CommandButton BotaoVerRoteiro 
         Caption         =   "Ver Roteiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   7875
         TabIndex        =   24
         ToolTipText     =   "Visualiza o Roteiro de Fabricação"
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         Caption         =   "Itens"
         Height          =   3975
         Index           =   3
         Left            =   225
         TabIndex        =   26
         Top             =   0
         Width           =   8865
         Begin VB.CommandButton BotaoAtualizar 
            Caption         =   "Atualizar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   5790
            TabIndex        =   61
            ToolTipText     =   "Abre a tela de Exportação de Projetos"
            Top             =   195
            Width           =   1230
         End
         Begin VB.TextBox Custeio 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   4470
            TabIndex        =   60
            Top             =   795
            Width           =   720
         End
         Begin VB.CommandButton BotaoCusteio 
            Caption         =   "Custeio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   105
            TabIndex        =   59
            ToolTipText     =   "Abre Browse dos Custeios de Roteiros cadastrados para o Produto"
            Top             =   3495
            Width           =   1155
         End
         Begin VB.CommandButton BotaoExportar 
            Caption         =   "Exportar ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   4515
            TabIndex        =   58
            ToolTipText     =   "Abre a tela de Exportação de Projetos"
            Top             =   195
            Width           =   1230
         End
         Begin VB.CommandButton BotaoDocDestino 
            Caption         =   "Abrir Destino ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   7065
            TabIndex        =   57
            ToolTipText     =   "Abre a tela do Destino dado ao item do Projeto"
            Top             =   195
            Width           =   1665
         End
         Begin VB.ComboBox DestinoPadrao 
            Height          =   315
            ItemData        =   "Projeto.ctx":0000
            Left            =   1590
            List            =   "Projeto.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   225
            Width           =   2325
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   315
            Left            =   6945
            TabIndex        =   53
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox PrecoItem 
            Height          =   315
            Left            =   5715
            TabIndex        =   54
            Top             =   1200
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
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
         Begin VB.TextBox Status 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   4455
            TabIndex        =   47
            Top             =   2490
            Width           =   3495
         End
         Begin VB.ComboBox Destino 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Projeto.ctx":0004
            Left            =   6105
            List            =   "Projeto.ctx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   2055
            Width           =   2070
         End
         Begin MSMask.MaskEdBox CustoTotal 
            Height          =   315
            Left            =   4890
            TabIndex        =   45
            Top             =   2070
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox CustoItem 
            Height          =   315
            Left            =   3660
            TabIndex        =   44
            Top             =   2070
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox DataMaxima 
            Height          =   315
            Left            =   2520
            TabIndex        =   43
            Top             =   2055
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataTermino 
            Height          =   315
            Left            =   1380
            TabIndex        =   42
            Top             =   2055
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataInicio 
            Height          =   315
            Left            =   255
            TabIndex        =   41
            Top             =   2055
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoProdItens 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2625
            TabIndex        =   37
            Top             =   1605
            Width           =   3660
         End
         Begin VB.ComboBox UMProdItens 
            Height          =   315
            Left            =   6285
            TabIndex        =   36
            Top             =   1605
            Width           =   885
         End
         Begin VB.ComboBox VersaoProdItens 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Projeto.ctx":0008
            Left            =   1740
            List            =   "Projeto.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1620
            Width           =   930
         End
         Begin MSMask.MaskEdBox ProdutoItens 
            Height          =   315
            Left            =   240
            TabIndex        =   38
            Top             =   1590
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeProdItens 
            Height          =   315
            Left            =   7170
            TabIndex        =   39
            Top             =   1605
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2565
            Left            =   90
            TabIndex        =   18
            Top             =   645
            Width           =   8670
            _ExtentX        =   15293
            _ExtentY        =   4524
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label Label1 
            Caption         =   "Destino Padrão:"
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
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   285
            Width           =   1410
         End
         Begin VB.Label CustoTotalProjeto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4395
            TabIndex        =   52
            Top             =   3540
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Custo Total:"
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
            Index           =   3
            Left            =   3285
            TabIndex        =   51
            Top             =   3570
            Width           =   1050
         End
         Begin VB.Label PrecoTotalProjeto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7260
            TabIndex        =   50
            Top             =   3540
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Preço Total:"
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
            Left            =   6135
            TabIndex        =   49
            Top             =   3570
            Width           =   1065
         End
      End
      Begin VB.CommandButton BotaoVersaoKitBase 
         Caption         =   "Versão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   1500
         TabIndex        =   21
         ToolTipText     =   "Abre o Browse de Versões do Kit para o Produto se este for Produzível"
         Top             =   4080
         Width           =   1155
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
         Height          =   355
         Left            =   2685
         TabIndex        =   22
         Top             =   4080
         Visible         =   0   'False
         Width           =   1050
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
         Height          =   355
         Left            =   225
         TabIndex        =   20
         ToolTipText     =   "Abre o Browse de Produtos"
         Top             =   4080
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4590
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   960
      Width           =   9225
      Begin VB.CommandButton BotaoGrafico 
         Caption         =   "Cronograma Gráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   180
         TabIndex        =   48
         ToolTipText     =   "Abre a tela do Cronograma Gráfico do Projeto"
         Top             =   4020
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Caption         =   "Outros"
         Height          =   1260
         Left            =   180
         TabIndex        =   29
         Top             =   2655
         Width           =   8865
         Begin VB.TextBox Observacao 
            Height          =   330
            Left            =   1725
            MaxLength       =   255
            TabIndex        =   9
            Top             =   750
            Width           =   6960
         End
         Begin MSComCtl2.UpDown UpDownDataCriacao 
            Height          =   300
            Left            =   7995
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataCriacao 
            Height          =   315
            Left            =   6840
            TabIndex        =   7
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Responsavel 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1740
            TabIndex        =   6
            Top             =   270
            Width           =   2805
         End
         Begin VB.Label Label3 
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
            Height          =   330
            Left            =   540
            TabIndex        =   34
            Top             =   810
            Width           =   1155
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data de criação:"
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
            Left            =   5310
            TabIndex        =   32
            Top             =   330
            Width           =   1440
         End
         Begin VB.Label PesponsavelLabel 
            AutoSize        =   -1  'True
            Caption         =   "Responsável:"
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
            Left            =   480
            TabIndex        =   30
            Top             =   330
            Width           =   1170
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   780
         Index           =   6
         Left            =   195
         TabIndex        =   25
         Top             =   1830
         Width           =   8865
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   6140
            TabIndex        =   5
            Top             =   285
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1710
            TabIndex        =   4
            Top             =   285
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
            Left            =   5535
            TabIndex        =   27
            Top             =   315
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
            Left            =   965
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   315
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1740
         Index           =   0
         Left            =   195
         TabIndex        =   19
         Top             =   30
         Width           =   8865
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2520
            Picture         =   "Projeto.ctx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   330
            Width           =   300
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1710
            TabIndex        =   0
            Top             =   330
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   315
            Left            =   1710
            TabIndex        =   3
            Top             =   1230
            Width           =   7020
            _ExtentX        =   12383
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReduzido 
            Height          =   315
            Left            =   1710
            TabIndex        =   2
            Top             =   765
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeReduzido 
            Caption         =   "Nome Reduzido:"
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
            Height          =   315
            Left            =   225
            TabIndex        =   40
            Top             =   795
            Width           =   1410
         End
         Begin VB.Label LabelObeservações 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            BeginProperty Font 
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
            TabIndex        =   33
            Top             =   1290
            Width           =   930
         End
         Begin VB.Label LabelCodigo 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   960
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   16
            Top             =   345
            Width           =   660
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7275
      ScaleHeight     =   450
      ScaleWidth      =   2055
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   90
      Width           =   2115
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1590
         Picture         =   "Projeto.ctx":00F6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1080
         Picture         =   "Projeto.ctx":0274
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   555
         Picture         =   "Projeto.ctx":07A6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   45
         Picture         =   "Projeto.ctx":0930
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5070
      Left            =   60
      TabIndex        =   14
      Top             =   570
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   8943
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
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
Attribute VB_Name = "ProjetoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Foi incluida a referência ao GlobaisPCP.dll em TelasFAT.vbp
'Por Jorge Specian - 17/05/2005

Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iLeCliente As Integer

Dim iFrameAtual As Integer

'Grid de Itens
Dim objGridItens As AdmGrid
Dim iGrid_ProdutoItens_Col As Integer
Dim iGrid_VersaoProdItens_Col As Integer
Dim iGrid_DescricaoProdItens_Col As Integer
Dim iGrid_UMProdItens_Col As Integer
Dim iGrid_QuantidadeProdItens_Col As Integer
Dim iGrid_CustoItem_Col As Integer
Dim iGrid_PrecoItem_Col As Integer
Dim iGrid_Custeio_Col As Integer
Dim iGrid_DataInicio_Col As Integer
Dim iGrid_DataTermino_Col As Integer
Dim iGrid_DataMaxima_Col As Integer
Dim iGrid_Destino_Col As Integer
Dim iGrid_Status_Col As Integer
Dim iGrid_CustoTotal_Col As Integer
Dim iGrid_PrecoTotal_Col As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCusteio As AdmEvento
Attribute objEventoCusteio.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Projeto"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Projeto"

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

Private Sub BotaoAtualizar_Click()

Dim lErro As Long
Dim objProjeto As ClassProjeto
Dim objProjetoItens As New ClassProjetoItens
   
On Error GoTo Erro_BotaoExportar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o código do Projeto não estiver preenchido ... Erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 134290
    
    Set objProjeto = New ClassProjeto
    
    objProjeto.lCodigo = StrParaLong(Trim(Codigo.Text))

    'Verifica se o Projeto existe, lendo no BD a partir do Código
    lErro = CF("Projeto_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> 139118 Then gError 134094
        
    If lErro <> SUCESSO Then gError 134095
    
    'Le os Itens do Projeto
    lErro = CF("Projeto_Le_Itens", objProjeto)
    If lErro <> SUCESSO And lErro <> 139126 Then gError 134096
    
    'Atualiza os dados da coleção de Projeto Itens na tela (GridItens)
    For Each objProjetoItens In objProjeto.colProjetoItens
                                        
        'Atualiza no Grid Itens
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_Destino_Col) = Seleciona_Destino(objProjetoItens.iDestino)

        lErro = Mostra_Status(objProjetoItens)
        If lErro <> SUCESSO Then gError 134200

    Next

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExportar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 134094, 134096, 134200
        
        Case 134095
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETO_NAO_CADASTRADO", gErr, objProjeto.lCodigo)
    
        Case 134290
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PROJETO_NAO_PREENCHIDO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165799)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoCusteio_Click()

Dim lErro As Long
Dim objCusteioRoteiro As New ClassCusteioRoteiro
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sFiltro As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoCusteio_Click

    'Se não tiver linha selecionada => Erro
    If GridItens.Row = 0 Then gError 136395
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)) = 0 Then gError 137428
    
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
    
    'formata o código do produto que está no grid
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134702

    'se o produto não existe cadastrado ...
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 134200

    sFiltro = "Produto = ? And Versao = ? And (DataValidade >= ? Or DataValidade = ?) "

    colSelecao.Add sProdutoFormatado
    colSelecao.Add sVersao
    If Len(DataCriacao.ClipText) <> 0 Then
        colSelecao.Add StrParaDate(DataCriacao.Text)
    Else
        colSelecao.Add gdtDataAtual
    End If
    colSelecao.Add DATA_NULA
    
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col)) <> 0 Then
        objCusteioRoteiro.lCodigo = StrParaLong(GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col))
    End If
    
    Call Chama_Tela("CusteioRoteirosLista", colSelecao, objCusteioRoteiro, objEventoCusteio, sFiltro)

    Exit Sub

Erro_BotaoCusteio_Click:

    Select Case gErr
    
        Case 134702

        Case 136395
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 137428
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165800)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDocDestino_Click()

Dim lErro As Long
Dim iDestino As Integer
Dim objProjetoItensRegGerados As New ClassProjetoItensRegGerados
   
On Error GoTo Erro_BotaoDocDestino_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se não tiver linha selecionada => Erro
    If GridItens.Row = 0 Then gError 136395
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)) = 0 Then gError 137428
   
    'Verifica se Destino está Preenchido
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Destino_Col)) = 0 Then gError 137429
    
    iDestino = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_Destino_Col))
    
    lErro = Verifica_Relacionamento(objProjetoItensRegGerados, iDestino)
    If lErro <> SUCESSO And lErro <> 137102 Then gError 137100
    
    If lErro <> SUCESSO Then gError 137102
    
    Select Case iDestino
   
        Case Is = ITEMDEST_ORCAMENTO_DE_VENDA
        
            'Chama a tela de orçamento de venda
            lErro = Abre_Tela_OrcamentoVenda(objProjetoItensRegGerados)
            If lErro <> SUCESSO Then gError 137101
        
        Case Is = ITEMDEST_PEDIDO_DE_VENDA
        
            'Chama a tela de pedido de venda
            lErro = Abre_Tela_PedidoVenda(objProjetoItensRegGerados)
            If lErro <> SUCESSO Then gError 137101
        
        Case Is = ITEMDEST_ORDEM_DE_PRODUCAO
        
            'Chama a tela de ordem de produção
        
        Case Is = ITEMDEST_NFISCAL_SIMPLES
        
            'Chama a tela de nota fiscal simples
        
        Case Is = ITEMDEST_NFISCAL_FATURA
            
            'Chama a tela de nota fiscal fatura
            
        Case Is = ITEMDEST_ORDEM_DE_SERVICO
            
            'Chama a tela de ordem de servico
            
    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoDocDestino_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 137100, 137101
            'erros tratados nas rotinas chamadas
            
        Case 137102
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_GERADO", gErr)

        Case 136395
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 137428
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 137429
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_PREENCHIDO", gErr, GridItens.Row)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165801)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjeto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 134290

    objProjeto.lCodigo = StrParaLong(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PROJETO", objProjeto.lCodigo)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui o Projeto
    lErro = CF("Projeto_Exclui", objProjeto)
    If lErro <> SUCESSO Then gError 134291

    'Limpa Tela
    Call Limpa_Tela_Projeto

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134290
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PROJETO_NAO_PREENCHIDO", gErr)

        Case 134291

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165802)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExportar_Click()

Dim lErro As Long
Dim objProjetoSeleciona As New ClassProjetoSeleciona
   
On Error GoTo Erro_BotaoExportar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass

    'Se o código do Projeto não estiver preenchido ... Erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 134290
    
    objProjetoSeleciona.lProjetoInicial = StrParaLong(Trim(Codigo.Text))
    objProjetoSeleciona.lProjetoFinal = StrParaLong(Trim(Codigo.Text))
    
    'Se tiver linha selecionada
    If GridItens.Row <> 0 Then
    
        objProjetoSeleciona.sProdutoInicial = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        objProjetoSeleciona.sProdutoFinal = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    
    End If
    
    'Chama a tela de exportação de projeto
    Call Chama_Tela("ExportarProjetos", objProjetoSeleciona)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExportar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 134290
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PROJETO_NAO_PREENCHIDO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165803)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165804)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGrafico_Click()

Dim lErro As Long
Dim objTelaGrafico As ClassTelaGrafico

On Error GoTo Erro_BotaoGrafico_Click:

    lErro = Atualiza_Cronograma(objTelaGrafico)
    If lErro <> SUCESSO Then gError 138246
    
    Call Chama_Tela_Nova_Instancia("TelaGrafico", objTelaGrafico)

    Exit Sub

Erro_BotaoGrafico_Click:

    Select Case gErr
    
        Case 138245, 138246

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165805)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Grava o Projeto
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 134101

    'Limpa Tela
    Call Limpa_Tela_Projeto
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    iClienteAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 134101

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165806)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134289
    
    Call Limpa_Tela_Projeto

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 134289

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165807)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_BotaoProdutos_Click
    
    If Me.ActiveControl Is ProdutoItens Then
    
        sProduto = Trim(ProdutoItens.Text)
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 134698
        
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134699
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    sFiltro = "Ativo = ? And Compras <> ?"
    
    colSelecao.Add PRODUTO_ATIVO
    colSelecao.Add PRODUTO_COMPRAVEL
        
    'Lista de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto, sFiltro)
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 134699
        
        Case 134698
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165808)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo numero disponível para um Projeto
    lErro = CF("Projeto_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 134339
    
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 134339
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165809)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoVerKit_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoVerKit

    If Me.ActiveControl Is ProdutoItens Then
    
        sProduto = Trim(ProdutoItens.Text)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    ElseIf Me.ActiveControl Is VersaoProdItens Then
    
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = Trim(VersaoProdItens.Text)
            
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 134698
        
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134756

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objKit.sProdutoRaiz = sProdutoFormatado
        If Len(sVersao) > 0 Then
        
            objKit.sVersao = sVersao
        
        Else
        
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 134202
        
        End If
            
        Call Chama_Tela("Kit", objKit)
    
    Else
         gError 134757
         
    End If

    Exit Sub
    
Erro_BotaoVerKit:

    Select Case gErr
    
        Case 134200, 134202, 134756
        
        Case 134698
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 134757
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO2", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165810)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoVerRoteiro_Click()

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoVerKit

    If Me.ActiveControl Is ProdutoItens Then
    
        sProduto = Trim(ProdutoItens.Text)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    ElseIf Me.ActiveControl Is VersaoProdItens Then
    
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = Trim(VersaoProdItens.Text)
            
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 134698
        
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134756

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
        If Len(sVersao) > 0 Then objRoteirosDeFabricacao.sVersao = sVersao
        
        Call Chama_Tela("RoteirosDeFabricacao", objRoteirosDeFabricacao)
    
    Else
         gError 134757
         
    End If

    Exit Sub
    
Erro_BotaoVerKit:

    Select Case gErr
            
        Case 134698
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 134756, 134200
        
        Case 134757
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO3", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165811)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoVersaoKitBase_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoVersaoKitBase_Click
    
    If Me.ActiveControl Is ProdutoItens Then
    
        sProduto = Trim(ProdutoItens.Text)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    ElseIf Me.ActiveControl Is VersaoProdItens Then
    
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = Trim(VersaoProdItens.Text)
            
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 134698
        
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134756

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objKit.sProdutoRaiz = sProdutoFormatado
        If Len(sVersao) > 0 Then objKit.sVersao = sVersao
            
        colSelecao.Add sProdutoFormatado
        
        Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersao)
    
    Else
         gError 134757
         
    End If

    Exit Sub

Erro_BotaoVersaoKitBase_Click:

    Select Case gErr

        Case 134698
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 134756
        
        Case 134757
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165812)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

    If iLeCliente <> iClienteAlterado Then

        Call Cliente_Preenche

    End If

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o cliente foi alterado
    If iClienteAlterado = 0 Then Exit Sub
    
    'Se cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
    
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 35713

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 35714

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        'Seleciona filial na Combo Filial
        If iCodFilial = FILIAL_MATRIZ Then
            Filial.ListIndex = 0
        Else
            Call CF("Filial_Seleciona", Filial, iCodFilial)
        End If
        
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        Filial.Clear

    End If

    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 35713, 35714, 35739

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165813)

    End Select

    Exit Sub

End Sub



Private Sub DestinoPadrao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DestinoPadrao_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iLinha As Integer

On Error GoTo Erro_DestinoPadrao_Click

    'se tem itens no grid e Destino Padrão preenchido ...
    If objGridItens.iLinhasExistentes > 0 And Len(DestinoPadrao.Text) <> 0 Then
    
        'Pergunta ao usuário se confirma a alteração de todos os Destinos
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_ALTERACAO_DOS_DESTINOS", DestinoPadrao.Text)
        
        'se SIM
        If vbMsgRes = vbYes Then
        
            'Altera cada linha do grid para o DestinoPadrao
            For iLinha = 1 To objGridItens.iLinhasExistentes
            
                GridItens.TextMatrix(iLinha, iGrid_Destino_Col) = DestinoPadrao.Text
            
            Next
            
        End If
        
    End If
    
    Exit Sub

Erro_DestinoPadrao_Click:

    Select Case gErr

        Case 136475, 136496

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165814)

    End Select
    
    Exit Sub

End Sub


Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iDestino As Integer
Dim objProjetoItensRegGerados As New ClassProjetoItensRegGerados

On Error GoTo Erro_GridItens_KeyDown
    
    iDestino = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_Destino_Col))
    
    lErro = Verifica_Relacionamento(objProjetoItensRegGerados, iDestino)
    If lErro <> SUCESSO And lErro <> 137102 Then gError 137100
    
    If lErro = SUCESSO Then gError 137101

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    
    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 137101
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPROJETO_NAO_PODE_EXCLUIR", gErr, GridItens.Row)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165815)
    
    End Select

    Exit Sub
        
End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Cliente.Text)) > 0 Then objCliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjeto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objProjeto.lCodigo = StrParaLong(Trim(Codigo.Text))

    End If

    Call Chama_Tela("ProjetoLista", colSelecao, objProjeto, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165816)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCliente As ClassCliente
Dim bCancel As Boolean

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165817)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjeto
Dim bCancel As Boolean

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objProjeto = obj1

    'Preenche campo Codigo
    Codigo.Text = objProjeto.lCodigo

    lErro = Traz_Projeto_Tela(objProjeto)
    If lErro <> SUCESSO And lErro <> 134095 Then gError 134292

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 134292

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165818)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCusteio_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCusteioRoteiro As New ClassCusteioRoteiro
Dim dCustoTotal As Double
Dim dPrecoTotal As Double
Dim dCustoItem As Double
Dim dPrecoItem As Double

On Error GoTo Erro_objEventoCusteio_evSelecao

    Set objCusteioRoteiro = obj1
    
    'Pega o Custo e Preço Total do Item que foi calculado no Custeio
    dCustoTotal = objCusteioRoteiro.dCustoTotalInsumosKit + objCusteioRoteiro.dCustoTotalInsumosMaq + objCusteioRoteiro.dCustoTotalMaoDeObra
    dPrecoTotal = objCusteioRoteiro.dPrecoTotalRoteiro
    
    'Calcula o Custo e Preco Unitário, dividindo pela quantidade utilizada para cálculo no Custeio
    '... é a quantidade que está gravada no BD para o CusteioRoteiro (não a do Grid)
    If objCusteioRoteiro.dQuantidade > 0 Then
        dCustoItem = dCustoTotal / objCusteioRoteiro.dQuantidade
        dPrecoItem = dPrecoTotal / objCusteioRoteiro.dQuantidade
    End If
    
    'Exibe os valores encontrados no Grid
    GridItens.TextMatrix(GridItens.Row, iGrid_CustoItem_Col) = Format(dCustoItem, "Standard")
    GridItens.TextMatrix(GridItens.Row, iGrid_PrecoItem_Col) = Format(dPrecoItem, "Standard")
    
    'Se já tem quantidade no Grid... Exibe os Totais
    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_QuantidadeProdItens_Col)) <> 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_CustoTotal_Col) = Format(dCustoItem * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantidadeProdItens_Col)), "Standard")
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dPrecoItem * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantidadeProdItens_Col)), "Standard")
    End If
    
    GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col) = objCusteioRoteiro.lCodigo
    
    'e, exibe o cálculo do total geral
    Call Calcula_CustoTotal
    Call Calcula_PrecoTotal
        
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub

Erro_objEventoCusteio_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165819)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 134708

    'Verifica se há algum produto repetido no grid
    For iLinha = 1 To objGridItens.iLinhasExistentes
        
        If iLinha <> GridItens.Row Then
                                                
            If GridItens.TextMatrix(iLinha, iGrid_ProdutoItens_Col) = sProdutoMascarado Then
                ProdutoItens.PromptInclude = False
                ProdutoItens.Text = ""
                ProdutoItens.PromptInclude = True
                gError 134709
                
            End If
                
        End If
                       
    Next
                
    GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col) = sProdutoMascarado
    
    Call Carrega_ComboVersoes(objProduto.sCodigo)
    
    If VersaoProdItens.ListCount > 0 Then
        
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)) = 0 Then
            Call VersaoProd_SelecionaPadrao(objProduto.sCodigo)
            GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col) = VersaoProdItens.Text
        End If
        
    End If

    GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoProdItens_Col) = objProduto.sDescricao
    
    Set objProdutoKit = New ClassProdutoKit
    
    objProdutoKit.sProdutoRaiz = objProduto.sCodigo
    objProdutoKit.sVersao = VersaoProdItens.Text
    
    'Lê o Produto Raiz do Kit para pegar seus dados
    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
    If lErro <> SUCESSO And lErro <> 34875 Then gError 134747
    
    'encontrou o ProdutoKit
    If lErro = SUCESSO Then
        
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col)) = 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col) = objProdutoKit.sUnidadeMed
        End If

    Else

        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col)) = 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col) = objProduto.sSiglaUMEstoque
        End If

    End If

    GridItens.TextMatrix(GridItens.Row, iGrid_Status_Col) = STRING_STATUS_NAO_EXPORTADO

    'verifica se precisa preencher o grid com uma nova linha
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 1347018, 134710
        
        Case 134709
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr, sProdutoMascarado, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165820)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1

    Call VersaoProd_Seleciona(objKit.sVersao)
    
    GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col) = VersaoProdItens.Text
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165821)

    End Select

    Exit Sub
    
End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If

End Sub

Private Sub UpDownDataCriacao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCriacao_DownClick

    DataCriacao.SetFocus

    If Len(DataCriacao.ClipText) > 0 Then

        sData = DataCriacao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 134750

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_DownClick:

    Select Case gErr

        Case 134750

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165822)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCriacao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCriacao_UpClick

    DataCriacao.SetFocus

    If Len(Trim(DataCriacao.ClipText)) > 0 Then

        sData = DataCriacao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 134751

        DataCriacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCriacao_UpClick:

    Select Case gErr

        Case 134751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165823)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
        
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
        
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
    
Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    Call ComandoSeta_Liberar(Me.Name)

    Set objEventoCliente = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVersao = Nothing
    Set objEventoCodigo = Nothing
    Set objEventoCusteio = Nothing
    
    Set objGridItens = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165824)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
   
    iFrameAtual = 1
    
    Set objEventoCliente = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Set objEventoCodigo = New AdmEvento
    Set objEventoCusteio = New AdmEvento
    
    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    Responsavel.Caption = gsUsuario
        
    'Grid Itens
    Set objGridItens = New AdmGrid
    
    'tela em questão
    Set objGridItens.objForm = Me
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 134340
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoItens)
    If lErro <> SUCESSO Then gError 134078
    
    lErro = CarregaComboDestino(DestinoPadrao)
    If lErro <> SUCESSO Then gError 134077
    
    lErro = CarregaComboDestino(Destino)
    If lErro <> SUCESSO Then gError 134077
    
    Call Calcula_CustoTotal
    Call Calcula_PrecoTotal
    
    iAlterado = 0
    iClienteAlterado = 0
    iLeCliente = 0
            
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 134340, 134200
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165825)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objProjeto As ClassProjeto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objProjeto Is Nothing) Then

        lErro = Traz_Projeto_Tela(objProjeto)
        If lErro <> SUCESSO And lErro <> 134095 And lErro <> 134097 Then gError 134280
        
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165826)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_ProdutoItens_Col

                    lErro = Saida_Celula_ProdutoItens(objGridInt)
                    If lErro <> SUCESSO Then gError 134371

                Case iGrid_VersaoProdItens_Col

                    lErro = Saida_Celula_VersaoProdItens(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
                
                Case iGrid_UMProdItens_Col

                    lErro = Saida_Celula_UMProdItens(objGridInt)
                    If lErro <> SUCESSO Then gError 134373
                
                Case iGrid_QuantidadeProdItens_Col

                    lErro = Saida_Celula_QuantidadeProdItens(objGridInt)
                    If lErro <> SUCESSO Then gError 134374
        
                Case iGrid_DataInicio_Col

                    lErro = Saida_Celula_DataInicio(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
                
                Case iGrid_DataTermino_Col

                    lErro = Saida_Celula_DataTermino(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
        
                Case iGrid_DataMaxima_Col

                    lErro = Saida_Celula_DataMaxima(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
        
                Case iGrid_PrecoItem_Col

                    lErro = Saida_Celula_PrecoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
        
                Case iGrid_CustoItem_Col

                    lErro = Saida_Celula_CustoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
        
                Case iGrid_Destino_Col

                    lErro = Saida_Celula_Destino(objGridInt)
                    If lErro <> SUCESSO Then gError 134372
        
            End Select
        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 134375

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 134370 To 134374
            'erros tratatos nas rotinas chamadas
        
        Case 134375
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165827)

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

Private Sub GridItens_GotFocus()
    
    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_LeaveCell()
    
    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)
    
End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Custo Unitário")
    objGrid.colColuna.Add ("Preço Unitário")
    objGrid.colColuna.Add ("Custeio")
    objGrid.colColuna.Add ("Data Início")
    objGrid.colColuna.Add ("Data Término")
    objGrid.colColuna.Add ("Data Máxima")
    objGrid.colColuna.Add ("Destino")
    objGrid.colColuna.Add ("Status")
    objGrid.colColuna.Add ("Custo Total")
    objGrid.colColuna.Add ("Preço Total")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ProdutoItens.Name)
    objGrid.colCampo.Add (VersaoProdItens.Name)
    objGrid.colCampo.Add (DescricaoProdItens.Name)
    objGrid.colCampo.Add (UMProdItens.Name)
    objGrid.colCampo.Add (QuantidadeProdItens.Name)
    objGrid.colCampo.Add (CustoItem.Name)
    objGrid.colCampo.Add (PrecoItem.Name)
    objGrid.colCampo.Add (Custeio.Name)
    objGrid.colCampo.Add (DataInicio.Name)
    objGrid.colCampo.Add (DataTermino.Name)
    objGrid.colCampo.Add (DataMaxima.Name)
    objGrid.colCampo.Add (Destino.Name)
    objGrid.colCampo.Add (Status.Name)
    objGrid.colCampo.Add (CustoTotal.Name)
    objGrid.colCampo.Add (PrecoTotal.Name)

    'Colunas do Grid
    iGrid_ProdutoItens_Col = 1
    iGrid_VersaoProdItens_Col = 2
    iGrid_DescricaoProdItens_Col = 3
    iGrid_UMProdItens_Col = 4
    iGrid_QuantidadeProdItens_Col = 5
    iGrid_CustoItem_Col = 6
    iGrid_PrecoItem_Col = 7
    iGrid_Custeio_Col = 8
    iGrid_DataInicio_Col = 9
    iGrid_DataTermino_Col = 10
    iGrid_DataMaxima_Col = 11
    iGrid_Destino_Col = 12
    iGrid_Status_Col = 13
    iGrid_CustoTotal_Col = 14
    iGrid_PrecoTotal_Col = 15

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sProduto As String
Dim objProdutos As ClassProduto
Dim objClasseUM As ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Guardo o valor do Codigo do Produto
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134767
    
    'Grid Itens
    If objControl.Name = "ProdutoItens" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = False

        Else
            objControl.Enabled = True
        
        End If

    ElseIf objControl.Name = "VersaoProdItens" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True
            Call Carrega_ComboVersoes(sProdutoFormatado)
            Call VersaoProd_Seleciona(GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col))

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "UMProdItens" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objControl.Enabled = True

            Set objProdutos = New ClassProduto

            objProdutos.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProdutos)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 134768

            Set objClasseUM = New ClassClasseUM
            
            objClasseUM.iClasse = objProdutos.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError 134769

            'Se tem algum valor para UMProdItens do Grid
            If Len(GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col)) > 0 Then
                'Guardo o valor da UMProdItens da Linha
                sUnidadeMed = GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col)
            Else
                'Senão coloco o do Produto em estoque
                sUnidadeMed = objProdutos.sSiglaUMEstoque
            End If
            
            'Limpar as Unidades utilizadas anteriormente
            UMProdItens.Clear

            For Each objUnidadeDeMedida In colSiglas
                UMProdItens.AddItem objUnidadeDeMedida.sSigla
            Next

            UMProdItens.AddItem ""

            'Tento selecionar na Combo a Unidade anterior
            If UMProdItens.ListCount <> 0 Then

                For iIndice = 0 To UMProdItens.ListCount - 1

                    If UMProdItens.List(iIndice) = sUnidadeMed Then
                        UMProdItens.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If
            
        Else
            objControl.Enabled = False
        
        End If

    ElseIf objControl.Name = "QuantidadeProdItens" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "DescricaoProdItens" Then

        objControl.Enabled = False
        
    ElseIf objControl.Name = "DataInicio" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "DataTermino" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "DataMaxima" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "CustoItem" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "CustoTotal" Then
        
        objControl.Enabled = False
        
    ElseIf objControl.Name = "PrecoItem" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
        
    ElseIf objControl.Name = "PrecoTotal" Then
        
        objControl.Enabled = False
                
    ElseIf objControl.Name = "Custeio" Then
        
        objControl.Enabled = False
                
    ElseIf objControl.Name = "Destino" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If
            
    ElseIf objControl.Name = "Status" Then
        
        objControl.Enabled = False
        
    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 134767 To 134769

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165828)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ProdutoItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ProdutoItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ProdutoItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescricaoProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescricaoProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VersaoProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VersaoProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub VersaoProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub VersaoProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = VersaoProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub UMProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub UMProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UMProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub QuantidadeProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub QuantidadeProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantidadeProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ProdutoItens(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ProdutoItens do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProdutoAnterior As String
Dim sCodProduto As String
Dim sVersao As String
Dim iLinha As Integer
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_Saida_Celula_ProdutoItens

    Set objGridInt.objControle = ProdutoItens

    sProdutoAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    sCodProduto = ProdutoItens.Text
        
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134771
    
    'Se o campo foi preenchido
    If Len(sProdutoFormatado) > 0 And sCodProduto <> sProdutoAnterior Then

        lErro = Mascara_RetornaProdutoTela(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134772
        
        'Verifica se há algum produto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridItens.Row Then
                                                    
                If GridItens.TextMatrix(iLinha, iGrid_ProdutoItens_Col) = sProdutoMascarado Then
                    ProdutoItens.PromptInclude = False
                    ProdutoItens.Text = ""
                    ProdutoItens.PromptInclude = True
                    gError 134773
                    
                End If
                    
            End If
                           
        Next
        
        Set objProdutos = New ClassProduto

        objProdutos.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134774
        
        'Verifica se o produto pode compor um Kit
        If objProdutos.iAtivo <> 0 And objProdutos.iGerencial <> 0 And _
            objProdutos.iKitBasico <> 1 And objProdutos.iKitInt <> 1 Then gError 134775
        
        Call Carrega_ComboVersoes(objProdutos.sCodigo)
        
        If VersaoProdItens.ListCount > 0 Then
            
            If Len(GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)) = 0 Then
                Call VersaoProd_SelecionaPadrao(objProdutos.sCodigo)
                GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col) = VersaoProdItens.Text
            End If
            
        End If
        
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)

        GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoProdItens_Col) = objProdutos.sDescricao
        
        Set objProdutoKit = New ClassProdutoKit
        
        objProdutoKit.sProdutoRaiz = sProdutoFormatado
        objProdutoKit.sVersao = sVersao
        
        'Lê o Produto Raiz do Kit para pegar seus dados
        lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
        If lErro <> SUCESSO And lErro <> 34875 Then gError 134747
        
        'encontrou o ProdutoKit
        If lErro = SUCESSO Then
            
            If Len(GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col)) = 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col) = objProdutoKit.sUnidadeMed
            End If
    
        Else
    
            If Len(GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col)) = 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col) = objProdutos.sSiglaUMEstoque
            End If
    
        End If
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Status_Col) = STRING_STATUS_NAO_EXPORTADO
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
            
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_ProdutoItens = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoItens:

    Saida_Celula_ProdutoItens = gErr

    Select Case gErr
        
        Case 134200
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165829)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VersaoProdItens(objGridInt As AdmGrid) As Long
'Faz a crítica da célula VersaoProdItens do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_VersaoProdItens

    Set objGridInt.objControle = VersaoProdItens

    GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col) = VersaoProdItens.Text
    
    'Se o campo foi preenchido
    If Len(Trim(VersaoProdItens.Text)) > 0 Then
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_VersaoProdItens = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoProdItens:

    Saida_Celula_VersaoProdItens = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165830)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UMProdItens(objGridInt As AdmGrid) As Long
'Faz a crítica da célula UMProdItens do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_UMProdItens

    Set objGridInt.objControle = UMProdItens

    'Se o campo foi preenchido
    If Len(Trim(UMProdItens.Text)) > 0 Then

        GridItens.TextMatrix(GridItens.Row, iGrid_UMProdItens_Col) = UMProdItens.Text

        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_UMProdItens = SUCESSO

    Exit Function

Erro_Saida_Celula_UMProdItens:

    Saida_Celula_UMProdItens = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165831)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeProdItens(objGridInt As AdmGrid) As Long
'Faz a crítica da célula QuantidadeProdItens do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidadeAtual As Double
Dim dCustoTotal As Double
Dim dPrecoTotal As Double

On Error GoTo Erro_Saida_Celula_QuantidadeProdItens

    Set objGridInt.objControle = QuantidadeProdItens

    'Se o campo foi preenchido
    If Len(Trim(QuantidadeProdItens.ClipText)) <> 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(QuantidadeProdItens.Text)
        If lErro <> SUCESSO Then gError 134393
        
        QuantidadeProdItens.Text = Formata_Estoque(QuantidadeProdItens.Text)
        
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_CustoItem_Col)) <> 0 Then
        
            dCustoTotal = StrParaDbl(QuantidadeProdItens.Text) * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_CustoItem_Col))
            GridItens.TextMatrix(GridItens.Row, iGrid_CustoTotal_Col) = Format(dCustoTotal, "Standard")
        
        End If
        
        Call Calcula_CustoTotal
            
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoItem_Col)) <> 0 Then
        
            dPrecoTotal = StrParaDbl(QuantidadeProdItens.Text) * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoItem_Col))
            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, "Standard")
        
        End If
            
        Call Calcula_PrecoTotal
            
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    Else
    
        GridItens.TextMatrix(GridItens.Row, iGrid_CustoTotal_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = ""
        
        Call Calcula_CustoTotal
        Call Calcula_PrecoTotal
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_QuantidadeProdItens = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeProdItens:

    Saida_Celula_QuantidadeProdItens = gErr

    Select Case gErr
    
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165832)

    End Select

    Exit Function

End Function

Private Sub Carrega_ComboVersoes(ByVal sProdutoRaiz As String)
    
Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
    
On Error GoTo Erro_Carrega_ComboVersoes
    
    VersaoProdItens.Enabled = True
    
    'Limpa a Combo
    VersaoProdItens.Clear
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProdutoRaiz
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 134805
    
    VersaoProdItens.AddItem ""
    
    'Carrega a Combo com os Dados da Colecao
    For Each objKit In colKits
    
        VersaoProdItens.AddItem (objKit.sVersao)
                       
    Next
    
    Exit Sub
    
Erro_Carrega_ComboVersoes:

    Select Case gErr
    
        Case 134805
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165833)
    
    End Select
    
End Sub

Private Function VersaoProd_SelecionaPadrao(sProduto As String)

Dim lErro As Long
Dim objKit As New ClassKit
    
On Error GoTo Erro_VersaoProd_SelecionaPadrao
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProduto
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError 134806
        
    Call VersaoProd_Seleciona(objKit.sVersao)
    
    VersaoProd_SelecionaPadrao = SUCESSO
    
    Exit Function

Erro_VersaoProd_SelecionaPadrao:

    VersaoProd_SelecionaPadrao = gErr
    
    Select Case gErr

        Case 134806
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165834)

    End Select

    Exit Function
    
End Function

Private Sub VersaoProd_Seleciona(sVersao As String)
Dim iIndice As Integer

    VersaoProdItens.ListIndex = -1
    For iIndice = 0 To VersaoProdItens.ListCount - 1
        If VersaoProdItens.List(iIndice) = sVersao Then
            VersaoProdItens.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

'por Jorge Specian - chamada por Cliente_Change para localizar pela parte digitada do Nome
'Reduzido do Cliente através da CF Cliente_Pesquisa_NomeReduzido em RotinasCRFAT.ClassCRFATSelect
Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134014

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134014

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165835)

    End Select
    
    Exit Sub

End Sub

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjeto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Projeto"

    'Lê os dados da Tela Projeto
    lErro = Move_Tela_Memoria(objProjeto)
    If lErro <> SUCESSO Then gError 134722

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objProjeto.lCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 134722

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165836)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjeto

On Error GoTo Erro_Tela_Preenche

    objProjeto.lCodigo = colCampoValor.Item("Codigo").vValor

    If objProjeto.lCodigo <> 0 Then
        lErro = Traz_Projeto_Tela(objProjeto)
        If lErro <> SUCESSO Then gError 134723
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 134723

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165837)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Projeto() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Projeto
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    DestinoPadrao.ListIndex = 0

    DataCriacao.PromptInclude = False
    DataCriacao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCriacao.PromptInclude = True
    
    Responsavel.Caption = gsUsuario
        
    Call Grid_Limpa(objGridItens)
    
    Filial.Clear
    
    iAlterado = 0
    iClienteAlterado = 0
    iLeCliente = 0
    
    Call Calcula_CustoTotal
    Call Calcula_PrecoTotal
    
    Limpa_Tela_Projeto = SUCESSO

    Exit Function

Erro_Limpa_Tela_Projeto:

    Limpa_Tela_Projeto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165838)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objProjeto As ClassProjeto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProjetoItens As ClassProjetoItens
Dim objCliente As New ClassCliente
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCusteioRoteiro As ClassCusteioRoteiro

On Error GoTo Erro_Move_Tela_Memoria

    objProjeto.lCodigo = StrParaLong(Codigo.Text)
    objProjeto.sNomeReduzido = NomeReduzido.Text
    objProjeto.sDescricao = Descricao.Text
    
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 135701
        
        'Não encontrou p Cliente --> erro
        If lErro = 12348 Then gError 135702

        objProjeto.lCodCliente = objCliente.lCodigo
        
    End If

    objProjeto.iCodFilial = Codigo_Extrai(Filial.Text)

    objProjeto.sResponsavel = Responsavel.Caption
    
    If Len(Trim(DataCriacao.ClipText)) > 0 Then
        objProjeto.dtDataCriacao = CDate(DataCriacao.Text)
    Else
        objProjeto.dtDataCriacao = DATA_NULA
    End If
        
    objProjeto.sObservacao = Observacao.Text
    
    'Ir preenchendo a colecao no objProjeto com todas as linhas "existentes" do grid Itens
    For iIndice = 1 To objGridItens.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ProdutoItens_Col))) = 0 Then Exit For
        
        Set objProdutos = New ClassProduto
        
        lErro = CF("Produto_Formata", Trim(GridItens.TextMatrix(iIndice, iGrid_ProdutoItens_Col)), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 134080
        
        objProdutos.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134081
        
        Set objProjetoItens = New ClassProjetoItens
        
        objProjetoItens.sProduto = objProdutos.sCodigo
        objProjetoItens.sVersao = GridItens.TextMatrix(iIndice, iGrid_VersaoProdItens_Col)
        objProjetoItens.iSeq = iIndice
        objProjetoItens.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantidadeProdItens_Col))
        objProjetoItens.sUMedida = GridItens.TextMatrix(iIndice, iGrid_UMProdItens_Col)
        objProjetoItens.dtDataInicioPrev = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col))
        objProjetoItens.dtDataTerminoPrev = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataTermino_Col))
        objProjetoItens.dtDataMaxTermino = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataMaxima_Col))
        objProjetoItens.dPrecoTotalItem = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoItem_Col))
        objProjetoItens.dCustoTotalItem = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_CustoItem_Col))
        objProjetoItens.iDestino = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Destino_Col))
        
        If Len(GridItens.TextMatrix(iIndice, iGrid_Custeio_Col)) <> 0 Then
            
            Set objCusteioRoteiro = New ClassCusteioRoteiro
            
            objCusteioRoteiro.lCodigo = StrParaLong(GridItens.TextMatrix(iIndice, iGrid_Custeio_Col))
            
            'Verifica se o CusteioRoteiro existe, lendo no BD
            lErro = CF("CusteioRoteiro_Le", objCusteioRoteiro)
            If lErro <> SUCESSO And lErro <> 134449 Then gError 134094
            
            objProjetoItens.lNumIntDocCusteio = objCusteioRoteiro.lNumIntDoc
        
        End If
        
        objProjeto.colProjetoItens.Add objProjetoItens
    
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 135701
        
        Case 135702
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
        
        Case 134080, 134081
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, Trim(GridItens.TextMatrix(iIndice, iGrid_ProdutoItens_Col)))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165839)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objProjeto As New ClassProjeto
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 134084
        
    'Verifica se o NomeReduzido está preenchida
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 134085
    
    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 134086

    'Verifica se a FilialCliente está preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 134087
    
    'Verifica se a Data de Criação está preenchida
    If Len(Trim(DataCriacao.ClipText)) = 0 Then gError 137127
    
    'Para cada Item
    For iIndice = 1 To objGridItens.iLinhasExistentes
                
        'Verifica se item do projeto sofreu alterações e se podia ser alterado
        lErro = Verifica_Alteracao_Item(iIndice)
        If lErro <> SUCESSO Then gError 137105
        
        'Verifica se a Unidade de Medida do Produto foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UMProdItens_Col))) = 0 Then gError 134089
        
        'Verifica se a Quantidade foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_QuantidadeProdItens_Col))) = 0 Then gError 134088
                
        'Verifica se o Custo Unitário foi informado
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CustoItem_Col))) = 0 Then gError 134103
        
        'Verifica se o Preço Unitário foi informado
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoItem_Col))) = 0 Then gError 134104
        
        'Verifica se a Data Inicio foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col))) = 0 Then gError 134100
        
        'Verifica se a Data Termino foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DataTermino_Col))) = 0 Then gError 134101
        
        'Verifica se a Data Máxima foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DataMaxima_Col))) = 0 Then gError 134102
        
    Next
            
    'Preenche o objProjeto
    lErro = Move_Tela_Memoria(objProjeto)
    If lErro <> SUCESSO Then gError 134091

    lErro = Trata_Alteracao(objProjeto, objProjeto.lCodigo)
    If lErro <> SUCESSO Then gError 134092

    'Grava o Projeto no Banco de Dados
    lErro = CF("Projeto_Grava", objProjeto)
    If lErro <> SUCESSO Then gError 134093
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134084
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 134085
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)
        
        Case 134086
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
                                    
        Case 134087
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 137127
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACRIACAO_NAO_PREENCHIDA", gErr)
        
        Case 134089
            Call Rotina_Erro(vbOKOnly, "ERRO_UMEDIDA_PRODUTO_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 134088
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)
                            
        Case 134100
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_PROJETO_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 134101
            Call Rotina_Erro(vbOKOnly, "ERRO_DATATERMINO_PROJETO_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 134102
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAMAXIMA_PROJETO_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 134103
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTOUNIT_PROJETO_NAO_PREENCHIDO", gErr, iIndice)
            
        Case 134104
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOUNIT_PROJETO_NAO_PREENCHIDO", gErr, iIndice)
                        
        Case 134091, 134092, 134093, 137105
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165840)

    End Select

    Exit Function

End Function

Function Traz_Projeto_Tela(objProjeto As ClassProjeto) As Long

Dim lErro As Long
Dim objProjetoItens As New ClassProjetoItens
Dim objProdutos As ClassProduto
Dim sProdutoMascarado As String
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objCusteioRoteiro As ClassCusteioRoteiro

On Error GoTo Erro_Traz_Projeto_Tela

    If objProjeto.lCodigo > 0 Then
        
        'Verifica se o Projeto existe, lendo no BD a partir do Código
        lErro = CF("Projeto_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> 139118 Then gError 134094
            
        If lErro <> SUCESSO Then gError 134095
            
    ElseIf Len(Trim(objProjeto.sNomeReduzido)) > 0 Then
                
        'Verifica se o Projeto existe, lendo no BD a partir do NomeReduzido
        lErro = CF("Projeto_Le_NomeReduzido", objProjeto)
        If lErro <> SUCESSO And lErro <> 139161 Then gError 134096
    
        If lErro <> SUCESSO Then gError 134097
    
    End If

    'Limpa a tela
    Call Limpa_Tela_Projeto
        
    If objProjeto.lCodigo <> 0 Then Codigo.Text = CStr(objProjeto.lCodigo)
    NomeReduzido.Text = objProjeto.sNomeReduzido
    Descricao.Text = objProjeto.sDescricao
    
    iLeCliente = REGISTRO_ALTERADO
    
    Cliente.Text = CStr(objProjeto.lCodCliente)
    
    iLeCliente = 0
    
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then gError 35713

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 35714

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)
    
    'Seleciona filial na Combo Filial
    If objProjeto.iCodFilial <> 0 Then
        Call CF("Filial_Seleciona", Filial, objProjeto.iCodFilial)
    Else
        If iCodFilial = FILIAL_MATRIZ Then
            Filial.ListIndex = 0
        Else
            Call CF("Filial_Seleciona", Filial, iCodFilial)
        End If
    End If
    
    Responsavel.Caption = objProjeto.sResponsavel
    
    If objProjeto.dtDataCriacao <> DATA_NULA Then
        DataCriacao.PromptInclude = False
        DataCriacao.Text = Format(objProjeto.dtDataCriacao, "dd/mm/yy")
        DataCriacao.PromptInclude = True
    End If
    
    Observacao.Text = objProjeto.sObservacao
    
    'Le os Itens do Projeto
    lErro = CF("Projeto_Le_Itens", objProjeto)
    If lErro <> SUCESSO And lErro <> 139126 Then gError 134096
    
    'Exibe os dados da coleção de Projeto Itens na tela (GridItens)
    For Each objProjetoItens In objProjeto.colProjetoItens
        
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objProjetoItens.sProduto
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134099
        
        lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134100
                                
        'Insere no Grid Itens
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_ProdutoItens_Col) = sProdutoMascarado
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_VersaoProdItens_Col) = objProjetoItens.sVersao
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_DescricaoProdItens_Col) = objProdutos.sDescricao
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_UMProdItens_Col) = objProjetoItens.sUMedida
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_QuantidadeProdItens_Col) = Formata_Estoque(objProjetoItens.dQuantidade)
        If objProjetoItens.dtDataInicioPrev <> DATA_NULA Then
            GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_DataInicio_Col) = Format(objProjetoItens.dtDataInicioPrev, "dd/mm/yyyy")
        End If
        If objProjetoItens.dtDataTerminoPrev <> DATA_NULA Then
            GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_DataTermino_Col) = Format(objProjetoItens.dtDataTerminoPrev, "dd/mm/yyyy")
        End If
        If objProjetoItens.dtDataMaxTermino <> DATA_NULA Then
            GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_DataMaxima_Col) = Format(objProjetoItens.dtDataMaxTermino, "dd/mm/yyyy")
        End If
        If objProjetoItens.dCustoTotalItem <> 0 Then
            GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_CustoItem_Col) = Format(objProjetoItens.dCustoTotalItem, "Standard")
            GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_CustoTotal_Col) = Format(objProjetoItens.dQuantidade * objProjetoItens.dCustoTotalItem, "Standard")
        End If
        If objProjetoItens.dPrecoTotalItem <> 0 Then
            GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_PrecoItem_Col) = Format(objProjetoItens.dPrecoTotalItem, "Standard")
            GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_PrecoTotal_Col) = Format(objProjetoItens.dQuantidade * objProjetoItens.dPrecoTotalItem, "Standard")
        End If
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_Destino_Col) = Seleciona_Destino(objProjetoItens.iDestino)

        If objProjetoItens.lNumIntDocCusteio <> 0 Then
            
            Set objCusteioRoteiro = New ClassCusteioRoteiro
            
            objCusteioRoteiro.lNumIntDoc = objProjetoItens.lNumIntDocCusteio
            
            'Verifica se o CusteioRoteiro existe, lendo no BD
            lErro = CF("CusteioRoteiro_Le", objCusteioRoteiro)
            If lErro <> SUCESSO And lErro <> 137940 Then gError 134094
            
            GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_Custeio_Col) = objCusteioRoteiro.lCodigo
        
        End If

        lErro = Mostra_Status(objProjetoItens)
        If lErro <> SUCESSO Then gError 134200

    Next

    objGridItens.iLinhasExistentes = objProjeto.colProjetoItens.Count
    
    Call Calcula_CustoTotal
    Call Calcula_PrecoTotal
    
    iAlterado = 0
    iClienteAlterado = 0
    
    Traz_Projeto_Tela = SUCESSO

    Exit Function

Erro_Traz_Projeto_Tela:

    Traz_Projeto_Tela = gErr

    Select Case gErr

        Case 134094, 134096, 134098, 134099, 134100, 134655, 134200
            'Erros tratados nas rotinas chamadas
        
        Case 134095, 134097 '134095 = Não encontrou por código; 134097 = Não encontrou por NomeReduzido
            'Erros tratados na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165841)

    End Select

    Exit Function

End Function

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

        'Critica a Codigo
        lErro = Long_Critica(Trim(Codigo.Text))
        If lErro <> SUCESSO Then gError 134366

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 134366
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165842)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProjeto As New ClassProjeto

On Error GoTo Erro_NomeReduzido_Validate

    'Verifica se NomeReduzido está preenchida
    If Len(Trim(NomeReduzido.Text)) <> 0 Then

        objProjeto.sNomeReduzido = Trim(NomeReduzido.Text)

        lErro = Traz_Projeto_Tela(objProjeto)
        If lErro <> SUCESSO And lErro <> 134095 And lErro <> 134097 Then gError 134106

    End If

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    Select Case gErr

        Case 134753, 134106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165843)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NomeReduzido, iAlterado)
    
End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Verifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then


    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case 134753

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165844)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Descricao, iAlterado)
    
End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
    
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Filial_Validate

    'Veifica se Filial está preenchida
    If Len(Trim(Filial.Text)) <> 0 Then


    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 134753

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165845)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataCriacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCriacao_Validate

    'Verifica se DataCriacao está preenchida
    If Len(Trim(DataCriacao.Text)) <> 0 Then

        lErro = Data_Critica(DataCriacao.Text)
        If lErro <> SUCESSO Then gError 134752

    End If

    Exit Sub

Erro_DataCriacao_Validate:

    Cancel = True

    Select Case gErr

        Case 134752

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165846)

    End Select

    Exit Sub

End Sub

Private Sub DataCriacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCriacao, iAlterado)
    
End Sub

Private Sub DataCriacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Observacao_Validate

    'Verifica se Observacao está preenchida
    If Len(Trim(Observacao.Text)) <> 0 Then


    End If

    Exit Sub

Erro_Observacao_Validate:

    Cancel = True

    Select Case gErr

        Case 134753

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165847)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function CarregaComboDestino(objCombo As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CarregaComboDestino

    objCombo.AddItem ""
    objCombo.ItemData(objCombo.NewIndex) = 0
    
    objCombo.AddItem ITEMDEST_ORCAMENTO_DE_VENDA & SEPARADOR & STRING_ITEMDEST_ORCAMENTO_DE_VENDA
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_ORCAMENTO_DE_VENDA
    
    objCombo.AddItem ITEMDEST_PEDIDO_DE_VENDA & SEPARADOR & STRING_ITEMDEST_PEDIDO_DE_VENDA
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_PEDIDO_DE_VENDA
    
    objCombo.AddItem ITEMDEST_ORDEM_DE_PRODUCAO & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_PRODUCAO
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_ORDEM_DE_PRODUCAO
    
    objCombo.AddItem ITEMDEST_NFISCAL_SIMPLES & SEPARADOR & STRING_ITEMDEST_NFISCAL_SIMPLES
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_NFISCAL_SIMPLES
    
    objCombo.AddItem ITEMDEST_NFISCAL_FATURA & SEPARADOR & STRING_ITEMDEST_NFISCAL_FATURA
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_NFISCAL_FATURA
    
    objCombo.AddItem ITEMDEST_ORDEM_DE_SERVICO & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_SERVICO
    objCombo.ItemData(objCombo.NewIndex) = ITEMDEST_ORDEM_DE_SERVICO
        
    CarregaComboDestino = SUCESSO

    Exit Function

Erro_CarregaComboDestino:

    CarregaComboDestino = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165848)

    End Select

    Exit Function

End Function

Private Sub DataInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataInicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataInicio
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataTermino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataTermino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataTermino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataTermino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataTermino
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataMaxima_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataMaxima_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataMaxima_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataMaxima_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataMaxima
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrecoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CustoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CustoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CustoItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Status_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Status_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Status_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Status_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Status
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Destino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Destino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Destino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Destino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Destino
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_DataInicio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataInicio do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_DataInicio

    Set objGridInt.objControle = DataInicio

    'Se o campo foi preenchido
    If Len(Trim(DataInicio.ClipText)) <> 0 Then
        
        lErro = Data_Critica(DataInicio.Text)
        If lErro <> SUCESSO Then gError 134393
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_DataTermino_Col))) <> 0 Then
        
            If StrParaDate(DataInicio.Text) > StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataTermino_Col)) Then gError 134394
        
        Else
        
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_DataMaxima_Col))) <> 0 Then
            
                If StrParaDate(DataInicio.Text) > StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataMaxima_Col)) Then gError 134395
            
            End If
        
        End If
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134396

    Saida_Celula_DataInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_DataInicio:

    Saida_Celula_DataInicio = gErr

    Select Case gErr
        
        Case 134394
            Call Rotina_Erro(vbOKOnly, "ERRO_DTINICIO_MAIOR_DTTERMINO_PROJ", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134395
            Call Rotina_Erro(vbOKOnly, "ERRO_DTINICIO_MAIOR_DTMAXIMA_PROJ", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134393, 134396
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165849)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_DataTermino(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataTermino do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_DataTermino

    Set objGridInt.objControle = DataTermino
    
    'Se o campo foi preenchido
    If Len(Trim(DataTermino.ClipText)) <> 0 Then

        lErro = Data_Critica(DataTermino.Text)
        If lErro <> SUCESSO Then gError 134393
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_DataInicio_Col))) <> 0 Then
        
            If StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataInicio_Col)) > StrParaDate(DataTermino.Text) Then gError 134394
        
        End If
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_DataMaxima_Col))) <> 0 Then
        
            If StrParaDate(DataTermino.Text) > StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataMaxima_Col)) Then gError 134395
        
        End If
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_DataTermino = SUCESSO

    Exit Function

Erro_Saida_Celula_DataTermino:

    Saida_Celula_DataTermino = gErr

    Select Case gErr
        
        Case 134394
            Call Rotina_Erro(vbOKOnly, "ERRO_DTINICIO_MAIOR_DTTERMINO_PROJ", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134395
            Call Rotina_Erro(vbOKOnly, "ERRO_DTTERMINO_MAIOR_DTMAXIMA_PROJ", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
        Case 134393, 134396
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165850)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataMaxima(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataMaxima do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_DataMaxima

    Set objGridInt.objControle = DataMaxima

    'Se o campo foi preenchido
    If Len(Trim(DataMaxima.ClipText)) <> 0 Then
        
        lErro = Data_Critica(DataMaxima.Text)
        If lErro <> SUCESSO Then gError 134393
    
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_DataTermino_Col))) <> 0 Then
        
            If StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataTermino_Col)) > StrParaDate(DataMaxima.Text) Then gError 134394
        
        Else
        
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_DataInicio_Col))) <> 0 Then
            
                If StrParaDate(GridItens.TextMatrix(GridItens.Row, iGrid_DataInicio_Col)) > StrParaDate(DataMaxima.Text) Then gError 134395
            
            End If
    
        End If
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_DataMaxima = SUCESSO

    Exit Function

Erro_Saida_Celula_DataMaxima:

    Saida_Celula_DataMaxima = gErr

    Select Case gErr
        
        Case 134394
            Call Rotina_Erro(vbOKOnly, "ERRO_DTTERMINO_MAIOR_DTMAXIMA_PROJ", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 134395
            Call Rotina_Erro(vbOKOnly, "ERRO_DTINICIO_MAIOR_DTMAXIMA_PROJ", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 134393, 134396
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165851)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula PrecoItem do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dPrecoTotal As Double

On Error GoTo Erro_Saida_Celula_PrecoItem

    Set objGridInt.objControle = PrecoItem

    'Se alterou o valor...
    If StrParaDbl(PrecoItem.Text) <> StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoItem_Col)) Then
        'se valor original foi trazido de custeio...
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col)) <> 0 Then
        
            'Informa e pergunta ao usuário se confirma a alteração
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_ALTERACAO_PRECO", GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col))
        
            If vbMsgRes = vbNo Then gError 139048
            
            GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col) = ""
        
        End If
    End If

    'Se o campo foi preenchido
    If Len(Trim(PrecoItem.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(PrecoItem.Text)
        If lErro <> SUCESSO Then gError 134393
        
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_QuantidadeProdItens_Col)) <> 0 Then
        
            dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantidadeProdItens_Col)) * StrParaDbl(PrecoItem.Text)
            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, "Standard")
        
        End If
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
    
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = ""
    
    End If
    
    Call Calcula_PrecoTotal
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_PrecoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoItem:

    Saida_Celula_PrecoItem = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 139048
            PrecoItem.Text = GridItens.TextMatrix(GridItens.Row, iGrid_PrecoItem_Col)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165852)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CustoItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CustoItem do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dCustoTotal As Double

On Error GoTo Erro_Saida_Celula_CustoItem

    Set objGridInt.objControle = CustoItem

    'Se alterou o valor...
    If StrParaDbl(CustoItem.Text) <> StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_CustoItem_Col)) Then
        'se valor original foi trazido de custeio...
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col)) <> 0 Then
        
            'Informa e pergunta ao usuário se confirma a alteração
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_ALTERACAO_CUSTO", GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col))
        
            If vbMsgRes = vbNo Then gError 139049
            
            GridItens.TextMatrix(GridItens.Row, iGrid_Custeio_Col) = ""
        
        End If
    End If
    
    'Se o campo foi preenchido
    If Len(Trim(CustoItem.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(CustoItem.Text)
        If lErro <> SUCESSO Then gError 134393
        
        If Len(GridItens.TextMatrix(GridItens.Row, iGrid_QuantidadeProdItens_Col)) <> 0 Then
        
            dCustoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantidadeProdItens_Col)) * StrParaDbl(CustoItem.Text)
            GridItens.TextMatrix(GridItens.Row, iGrid_CustoTotal_Col) = Format(dCustoTotal, "Standard")
        
        End If
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
    
        GridItens.TextMatrix(GridItens.Row, iGrid_CustoTotal_Col) = ""
    
    End If
    
    Call Calcula_CustoTotal
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_CustoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoItem:

    Saida_Celula_CustoItem = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 139049
            CustoItem.Text = GridItens.TextMatrix(GridItens.Row, iGrid_CustoItem_Col)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165853)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Destino(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Destino do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Destino

    Set objGridInt.objControle = Destino

    GridItens.TextMatrix(GridItens.Row, iGrid_Destino_Col) = Destino.Text
    
    'Se o campo foi preenchido
    If Len(Trim(Destino.Text)) > 0 Then
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_Destino = SUCESSO

    Exit Function

Erro_Saida_Celula_Destino:

    Saida_Celula_Destino = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165854)

    End Select

    Exit Function

End Function

Public Function Atualiza_Cronograma(objTelaGrafico As ClassTelaGrafico) As Long

Dim objTelaGraficoItem As New ClassTelaGraficoItens
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sNL As String
Dim lErro As Long
Dim bPrimeira As Boolean
Dim bUltima As Boolean
Dim iDestino As Integer
Dim sStatus As String
Dim objProjetoItensRegGerados As ClassProjetoItensRegGerados
Dim objItemOrcamento As ClassItemOV
Dim objOrcamentoVenda As ClassOrcamentoVenda
Dim objItemPedido As ClassItemPedido
Dim objPedidoDeVenda As ClassPedidoDeVenda

On Error GoTo Erro_Atualiza_Cronograma

    GL_objMDIForm.MousePointer = vbHourglass

    'Para cada Item
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        'Verifica se a Data Inicio foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col))) = 0 Then gError 134100
        
        'Verifica se a Data Termino foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DataTermino_Col))) = 0 Then gError 134101
        
    Next

    Set objTelaGrafico = New ClassTelaGrafico

    Set objTelaGrafico.objTela = Me
    
    objTelaGrafico.sNomeTela = "Cronograma do Projeto"
    objTelaGrafico.iTamanhoDia = 540
    objTelaGrafico.iModal = DESMARCADO
    objTelaGrafico.iAtualizaRetornoClick = DESMARCADO

    sNL = Chr(13) & Chr(10)
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        'posiciona o grid na mesma linha que se está fazendo
        GridItens.Row = iIndice
        
        'Pega o destino informado no grid
        iDestino = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Destino_Col))
        
        Set objProjetoItensRegGerados = New ClassProjetoItensRegGerados
        
        'Encontra o relacionamento para encontrar o NumIntDocDestino
        lErro = Verifica_Relacionamento(objProjetoItensRegGerados, iDestino)
        If lErro <> SUCESSO And lErro <> 137102 Then gError 137100
        
        If lErro = SUCESSO Then
            sStatus = STRING_STATUS_EXPORTADO
        Else
            sStatus = STRING_STATUS_NAO_EXPORTADO
        End If
    
        'inicializa objeto do item do gráfico
        Set objTelaGraficoItem = New ClassTelaGraficoItens
    
        objTelaGraficoItem.dtDataFim = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataTermino_Col))
        objTelaGraficoItem.dtDataInicio = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col))
        objTelaGraficoItem.sTextoExibicao = "Produto: " & GridItens.TextMatrix(iIndice, iGrid_ProdutoItens_Col) & sNL & "Versão: " & GridItens.TextMatrix(iIndice, iGrid_VersaoProdItens_Col) & sNL & "Quantidade: " & GridItens.TextMatrix(iIndice, iGrid_QuantidadeProdItens_Col) & " " & GridItens.TextMatrix(iIndice, iGrid_UMProdItens_Col) & sNL & "Data Início: " & GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col) & sNL & "Data Fim: " & GridItens.TextMatrix(iIndice, iGrid_DataTermino_Col) & sNL & "Destino: " & GridItens.TextMatrix(iIndice, iGrid_Destino_Col) & sNL & "Status: " & sStatus
        objTelaGraficoItem.sNome = GridItens.TextMatrix(iIndice, iGrid_ProdutoItens_Col)
        
        'Cada destino em uma cor no gráfico (soma 1 para não ir vazio)
        objTelaGraficoItem.iIndiceCor = iDestino + 1  '1 -> Sem Destino (vbBlue), 2 (vbGreen), 3 (vbYellow), 4 (vbMagenta), 5 (vbCyan), 6 (12177627 - Bege), 7 (12153344 - Azul Médio)
        
        'Para cada Destino... Passa objeto para abertura no clique do gráfico
        Select Case iDestino
        
            Case Is = ITEMDEST_ORCAMENTO_DE_VENDA
            
                Set objItemOrcamento = New ClassItemOV
    
                'Preenche o obj com dados para abrir a tela
                objItemOrcamento.lNumIntDoc = objProjetoItensRegGerados.lNumIntDocDestino
            
                'Le o Item do Orcamento de Venda pelo seu NumIntDoc
                lErro = CF("Projeto_Le_ItemOV", objItemOrcamento)
                If lErro <> SUCESSO And lErro <> 139122 Then gError 137100
            
                If lErro = SUCESSO Then
                
                    Set objOrcamentoVenda = New ClassOrcamentoVenda
                    objOrcamentoVenda.lCodigo = objItemOrcamento.lCodOrcamento
                    objOrcamentoVenda.iFilialEmpresa = objItemOrcamento.iFilialEmpresa
                    
                    objTelaGraficoItem.sNomeTela = "OrcamentoVenda"
                    objTelaGraficoItem.colobj.Add objOrcamentoVenda
                
                End If
            
            Case Is = ITEMDEST_PEDIDO_DE_VENDA
            
                Set objItemPedido = New ClassItemPedido
    
                'Preenche o obj com dados para abrir a tela
                objItemPedido.lNumIntDoc = objProjetoItensRegGerados.lNumIntDocDestino
            
                'Le o Item do Orcamento de Venda pelo seu NumIntDoc
                lErro = CF("Projeto_Le_ItemPV", objItemPedido)
                If lErro <> SUCESSO And lErro <> 139242 Then gError 137100
            
                If lErro = SUCESSO Then
                
                    Set objPedidoDeVenda = New ClassPedidoDeVenda
                    objPedidoDeVenda.lCodigo = objItemPedido.lCodPedido
                    objPedidoDeVenda.iFilialEmpresa = objItemPedido.iFilialEmpresa
                    
                    objTelaGraficoItem.sNomeTela = "PedidoVenda"
                    objTelaGraficoItem.colobj.Add objPedidoDeVenda
                
                End If
            
            Case Is = ITEMDEST_ORDEM_DE_PRODUCAO
                        
            Case Is = ITEMDEST_NFISCAL_SIMPLES
            
            Case Is = ITEMDEST_NFISCAL_FATURA
                
            Case Is = ITEMDEST_ORDEM_DE_SERVICO
                
        End Select
        
        bPrimeira = True
        bUltima = True

        For iIndice1 = 1 To objGridItens.iLinhasExistentes
            'Se existe outra com data de início menor ou igual desde que
            'o seq esteja depois no grid então o item corrente
            'não é o primeiro
            If (StrParaDate(GridItens.TextMatrix(iIndice1, iGrid_DataInicio_Col)) < StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col))) Or (StrParaDate(GridItens.TextMatrix(iIndice1, iGrid_DataInicio_Col)) = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col)) And iIndice1 < iIndice) Then
                bPrimeira = False
            End If
        Next
        
        For iIndice1 = 1 To objGridItens.iLinhasExistentes
            'Se existe outra com data de término maior ou igual desde que
            'o seq esteja depois no grid então o item corrente
            'não é o último
            If (StrParaDate(GridItens.TextMatrix(iIndice1, iGrid_DataTermino_Col)) > StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataTermino_Col))) Or (StrParaDate(GridItens.TextMatrix(iIndice1, iGrid_DataTermino_Col)) = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataTermino_Col)) And iIndice1 > iIndice) Then
                bUltima = False
            End If
        Next

        If bPrimeira Then

            If objGridItens.iLinhasExistentes = 1 Then
                objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_INICIO_E_FIM
            Else
                objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_INICIO
            End If
        
        Else

            If bUltima Then
            
                objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_FIM
            
            End If

        End If

        objTelaGrafico.colItens.Add objTelaGraficoItem
    
    Next
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Atualiza_Cronograma = SUCESSO

    Exit Function

Erro_Atualiza_Cronograma:

    GL_objMDIForm.MousePointer = vbDefault

    Atualiza_Cronograma = gErr

    Select Case gErr
    
        Case 134100
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_PROJETO_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 134101
            Call Rotina_Erro(vbOKOnly, "ERRO_DATATERMINO_PROJETO_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165855)

    End Select
    
    Exit Function
    
End Function

Private Function Seleciona_Destino(iDestino As Integer) As String

Dim sTexto As String

    Select Case iDestino
    
        Case Is = ITEMDEST_ORCAMENTO_DE_VENDA
            sTexto = ITEMDEST_ORCAMENTO_DE_VENDA & SEPARADOR & STRING_ITEMDEST_ORCAMENTO_DE_VENDA
    
        Case Is = ITEMDEST_PEDIDO_DE_VENDA
            sTexto = ITEMDEST_PEDIDO_DE_VENDA & SEPARADOR & STRING_ITEMDEST_PEDIDO_DE_VENDA
    
        Case Is = ITEMDEST_ORDEM_DE_PRODUCAO
            sTexto = ITEMDEST_ORDEM_DE_PRODUCAO & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_PRODUCAO
    
        Case Is = ITEMDEST_NFISCAL_SIMPLES
            sTexto = ITEMDEST_NFISCAL_SIMPLES & SEPARADOR & STRING_ITEMDEST_NFISCAL_SIMPLES
    
        Case Is = ITEMDEST_NFISCAL_FATURA
            sTexto = ITEMDEST_NFISCAL_FATURA & SEPARADOR & STRING_ITEMDEST_NFISCAL_FATURA
    
        Case Is = ITEMDEST_ORDEM_DE_SERVICO
            sTexto = ITEMDEST_ORDEM_DE_SERVICO & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_SERVICO
    
    End Select
    
    Seleciona_Destino = sTexto
    
End Function

Public Function Mostra_Status(objProjetoItens As ClassProjetoItens) As Long

Dim lErro As Long
Dim objProjetoItensRegGerados As ClassProjetoItensRegGerados
Dim sTextoStatus As String
Dim objItemOrcamento As ClassItemOV
Dim objItemPedido As ClassItemPedido

On Error GoTo Erro_Mostra_Status

    Set objProjetoItensRegGerados = New ClassProjetoItensRegGerados

    objProjetoItensRegGerados.lNumIntDocItemProj = objProjetoItens.lNumIntDoc
    objProjetoItensRegGerados.iDestino = objProjetoItens.iDestino
    
    'Le a tabela de relacionamento
    lErro = CF("Projeto_Le_ItensRegGerados", objProjetoItensRegGerados)
    If lErro <> SUCESSO And lErro <> 139157 Then gError 137101
    
    'Se não tem relacionamentos -> não foi exportado ainda
    If lErro <> SUCESSO Then
    
        sTextoStatus = STRING_STATUS_NAO_EXPORTADO
    
    Else
    
        sTextoStatus = STRING_STATUS_EXPORTADO
    
        'Para cada destino monta o texto do status
        Select Case objProjetoItens.iDestino
        
            Case Is = ITEMDEST_ORCAMENTO_DE_VENDA
            
                sTextoStatus = sTextoStatus & SEPARADOR & STRING_ITEMDEST_ORCAMENTO_DE_VENDA
                
                Set objItemOrcamento = New ClassItemOV

                objItemOrcamento.lNumIntDoc = objProjetoItensRegGerados.lNumIntDocDestino
            
                'Le o Item do Orcamento de Venda pelo seu NumIntDoc
                lErro = CF("Projeto_Le_ItemOV", objItemOrcamento)
                If lErro <> SUCESSO And lErro <> 139122 Then gError 137100
                
                sTextoStatus = sTextoStatus & " " & objItemOrcamento.lCodOrcamento
            
            Case Is = ITEMDEST_PEDIDO_DE_VENDA
            
                sTextoStatus = sTextoStatus & SEPARADOR & STRING_ITEMDEST_PEDIDO_DE_VENDA
                
                Set objItemPedido = New ClassItemPedido

                objItemPedido.lNumIntDoc = objProjetoItensRegGerados.lNumIntDocDestino
            
                'Le o Item do Pedido de Venda pelo seu NumIntDoc
                lErro = CF("Projeto_Le_ItemPV", objItemPedido)
                If lErro <> SUCESSO And lErro <> 139242 Then gError 137100
                
                sTextoStatus = sTextoStatus & " " & objItemPedido.lCodPedido
            
            Case Is = ITEMDEST_ORDEM_DE_PRODUCAO
            
                sTextoStatus = sTextoStatus & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_PRODUCAO
            
            Case Is = ITEMDEST_NFISCAL_SIMPLES
            
                sTextoStatus = sTextoStatus & SEPARADOR & STRING_ITEMDEST_NFISCAL_SIMPLES
            
            Case Is = ITEMDEST_NFISCAL_FATURA
                
                sTextoStatus = sTextoStatus & SEPARADOR & STRING_ITEMDEST_NFISCAL_FATURA
                
            Case Is = ITEMDEST_ORDEM_DE_SERVICO
                
                sTextoStatus = sTextoStatus & SEPARADOR & STRING_ITEMDEST_ORDEM_DE_SERVICO
                
        End Select
      
    End If
    
    'Preenche Status no Grid
    GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_Status_Col) = sTextoStatus
    
    Mostra_Status = SUCESSO

    Exit Function

Erro_Mostra_Status:

    Mostra_Status = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165856)

    End Select
    
    Exit Function
    
End Function

Public Function Verifica_Relacionamento(objProjetoItensRegGerados As ClassProjetoItensRegGerados, iDestino As Integer) As Long

Dim lErro As Long
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProjeto As New ClassProjeto
Dim objProjetoItens As New ClassProjetoItens

On Error GoTo Erro_Verifica_Relacionamento

    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134702

    'se o produto não existe cadastrado ...
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 134200
    
    objProjeto.lCodigo = StrParaLong(Codigo.Text)
    
    'Verifica se o Projeto existe, lendo no BD a partir do Código
    lErro = CF("Projeto_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> 139118 Then gError 134094
        
    If lErro <> SUCESSO Then gError 134095
    
    'Le os Itens do Projeto
    lErro = CF("Projeto_Le_Itens", objProjeto)
    If lErro <> SUCESSO And lErro <> 139126 Then gError 134096
    
    'Percorre os Itens do Projeto para achar o Produto e Versão
    For Each objProjetoItens In objProjeto.colProjetoItens
    
        If objProjetoItens.sProduto = sProdutoFormatado And objProjetoItens.sVersao = Trim(sVersao) Then
    
            Exit For
    
        End If
        
    Next
    
    objProjetoItensRegGerados.lNumIntDocItemProj = objProjetoItens.lNumIntDoc
    objProjetoItensRegGerados.iDestino = iDestino
    
    'Le a tabela de relacionamento
    lErro = CF("Projeto_Le_ItensRegGerados", objProjetoItensRegGerados)
    If lErro <> SUCESSO And lErro <> 139157 Then gError 137101
    
    'se não tem relacionamento ... erro
    If lErro <> SUCESSO Then gError 137102
    
    Verifica_Relacionamento = SUCESSO

    Exit Function

Erro_Verifica_Relacionamento:

    Verifica_Relacionamento = gErr

    Select Case gErr
        
        Case 137102
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165857)

    End Select

    Exit Function

End Function

Public Function Abre_Tela_OrcamentoVenda(objProjetoItensRegGerados As ClassProjetoItensRegGerados) As Long

Dim lErro As Long
Dim objItemOrcamento As New ClassItemOV
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Abre_Tela_OrcamentoVenda

    objItemOrcamento.lNumIntDoc = objProjetoItensRegGerados.lNumIntDocDestino

    'Le o Item do Orcamento de Venda pelo seu NumIntDoc
    lErro = CF("Projeto_Le_ItemOV", objItemOrcamento)
    If lErro <> SUCESSO And lErro <> 139122 Then gError 137100

    If lErro <> SUCESSO Then gError 137101
    
    objOrcamentoVenda.lCodigo = objItemOrcamento.lCodOrcamento
    objOrcamentoVenda.iFilialEmpresa = objItemOrcamento.iFilialEmpresa

    'Chama a tela de Orçamento de Venda
    Call Chama_Tela("OrcamentoVenda", objOrcamentoVenda)
    
    Abre_Tela_OrcamentoVenda = SUCESSO

    Exit Function

Erro_Abre_Tela_OrcamentoVenda:

    Abre_Tela_OrcamentoVenda = gErr

    Select Case gErr
    
        Case 137100
            'erro tratado na rotina chamada
            
        Case 137101
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_GERADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165858)

    End Select

    Exit Function

End Function

Private Sub PrecoTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoTotal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CustoTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CustoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CustoTotal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Calcula_CustoTotal()

Dim dTotal As Double
Dim iLinha As Integer

    dTotal = 0
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        dTotal = dTotal + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_CustoTotal_Col))
    
    Next
    
    CustoTotalProjeto.Caption = Format(dTotal, "Standard")

End Sub

Private Sub Calcula_PrecoTotal()

Dim dTotal As Double
Dim iLinha As Integer

    dTotal = 0
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        dTotal = dTotal + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
    
    Next
    
    PrecoTotalProjeto.Caption = Format(dTotal, "Standard")

End Sub

Private Function Verifica_Alteracao_Item(ByVal iIndice As Integer) As Long

Dim lErro As Long
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProjeto As New ClassProjeto
Dim objProjetoItens As New ClassProjetoItens
Dim objProjetoItensRegGerados As New ClassProjetoItensRegGerados

On Error GoTo Erro_Verifica_Alteracao_Item

    sProduto = GridItens.TextMatrix(iIndice, iGrid_ProdutoItens_Col)
    sVersao = GridItens.TextMatrix(iIndice, iGrid_VersaoProdItens_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134702

    objProjeto.lCodigo = StrParaLong(Codigo.Text)
    
    'Verifica se o Projeto existe, lendo no BD a partir do Código
    lErro = CF("Projeto_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> 139118 Then gError 134094
        
    'Le os Itens do Projeto
    lErro = CF("Projeto_Le_Itens", objProjeto)
    If lErro <> SUCESSO And lErro <> 139126 Then gError 134096
    
    'Percorre os Itens do Projeto para achar o Produto e Versão
    For Each objProjetoItens In objProjeto.colProjetoItens
    
        If objProjetoItens.sProduto = sProdutoFormatado And objProjetoItens.sVersao = Trim(sVersao) Then
    
            objProjetoItensRegGerados.lNumIntDocItemProj = objProjetoItens.lNumIntDoc
            objProjetoItensRegGerados.iDestino = objProjetoItens.iDestino
    
            Exit For
    
        End If
        
    Next
    
    'Le a tabela de relacionamento
    lErro = CF("Projeto_Le_ItensRegGerados", objProjetoItensRegGerados)
    If lErro <> SUCESSO And lErro <> 139157 Then gError 137101
    
    'se tem relacionamento
    If lErro = SUCESSO Then
    
        If objProjetoItens.sProduto <> sProdutoFormatado Or _
            objProjetoItens.sVersao <> Trim(sVersao) Or _
            objProjetoItens.sUMedida <> GridItens.TextMatrix(iIndice, iGrid_UMProdItens_Col) Or _
            objProjetoItens.dQuantidade <> StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantidadeProdItens_Col)) Or _
            objProjetoItens.dCustoTotalItem <> StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_CustoItem_Col)) Or _
            objProjetoItens.dPrecoTotalItem <> StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoItem_Col)) Or _
            objProjetoItens.dtDataInicioPrev <> StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataInicio_Col)) Or _
            objProjetoItens.dtDataTerminoPrev <> StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataTermino_Col)) Or _
            objProjetoItens.dtDataMaxTermino <> StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataMaxima_Col)) Or _
            objProjetoItens.iDestino <> Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Destino_Col)) Then gError 134102
            
    End If
    
    Verifica_Alteracao_Item = SUCESSO

    Exit Function

Erro_Verifica_Alteracao_Item:

    Verifica_Alteracao_Item = gErr

    Select Case gErr
    
        Case 137702, 134094, 134096, 137101
            'erros tratados nas rotinas chamadas
            
        Case 134102
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPROJETO_NAO_PODE_ALTERAR", gErr, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165859)

    End Select

    Exit Function

End Function

Public Function Abre_Tela_PedidoVenda(objProjetoItensRegGerados As ClassProjetoItensRegGerados) As Long

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Abre_Tela_PedidoVenda

    objItemPedido.lNumIntDoc = objProjetoItensRegGerados.lNumIntDocDestino

    'Le o Item do Pedido de Venda pelo seu NumIntDoc
    lErro = CF("Projeto_Le_ItemPV", objItemPedido)
    If lErro <> SUCESSO And lErro <> 139242 Then gError 137100

    If lErro <> SUCESSO Then gError 137101
    
    objPedidoDeVenda.lCodigo = objItemPedido.lCodPedido
    objPedidoDeVenda.iFilialEmpresa = objItemPedido.iFilialEmpresa

    'Chama a tela de Pedido de Venda
    Call Chama_Tela("PedidoVenda", objPedidoDeVenda)
    
    Abre_Tela_PedidoVenda = SUCESSO

    Exit Function

Erro_Abre_Tela_PedidoVenda:

    Abre_Tela_PedidoVenda = gErr

    Select Case gErr
    
        Case 137100
            'erro tratado na rotina chamada
            
        Case 137101
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_GERADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165860)

    End Select

    Exit Function

End Function



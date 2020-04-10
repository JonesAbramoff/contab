VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CustoEmbMPOcx 
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   10110
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6150
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   795
      Width           =   9645
      Begin VB.Frame Frame9 
         Caption         =   "Material Comprado no mercado nacional com cotação em moeda"
         Height          =   675
         Left            =   60
         TabIndex        =   64
         Top             =   2730
         Width           =   9540
         Begin VB.ComboBox Moeda 
            Height          =   315
            ItemData        =   "CustoEmbMP.ctx":0000
            Left            =   3675
            List            =   "CustoEmbMP.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   225
            Width           =   1380
         End
         Begin MSMask.MaskEdBox CustoMoeda 
            Height          =   315
            Left            =   1710
            TabIndex        =   29
            Top             =   225
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cotacao 
            Height          =   315
            Left            =   5895
            TabIndex        =   31
            Top             =   225
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Moeda:"
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
            Height          =   210
            Left            =   3030
            TabIndex        =   79
            Top             =   285
            Width           =   615
         End
         Begin VB.Label LabelMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3210
            TabIndex        =   69
            Top             =   -105
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Custo Calculado:"
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
            Left            =   6825
            TabIndex        =   68
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label LabelCustoCalculado2 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8340
            TabIndex        =   67
            Top             =   225
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cotação:"
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
            Left            =   5100
            TabIndex        =   66
            Top             =   285
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Custo em Moeda:"
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
            Left            =   150
            TabIndex        =   65
            Top             =   285
            Width           =   1485
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informações Adicionais"
         Height          =   2715
         Left            =   60
         TabIndex        =   22
         Top             =   3405
         Width           =   9540
         Begin VB.CommandButton BotaoConsultaCusto 
            Caption         =   "Consultar Custos Cadastrados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   6720
            TabIndex        =   56
            Top             =   1875
            Width           =   2640
         End
         Begin VB.TextBox TaxaInfo 
            Height          =   315
            Left            =   3645
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   765
            Width           =   780
         End
         Begin VB.TextBox ICMSInfo 
            Height          =   315
            Left            =   2745
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   765
            Width           =   780
         End
         Begin VB.TextBox FreteInfo 
            Height          =   315
            Left            =   1785
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   510
            Width           =   840
         End
         Begin VB.TextBox MoedaInfo 
            Height          =   315
            Left            =   1665
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   120
            Width           =   870
         End
         Begin VB.TextBox DataInfo 
            Height          =   315
            Left            =   450
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   750
            Width           =   1050
         End
         Begin VB.TextBox CustoInfo 
            Height          =   345
            Left            =   330
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   360
            Width           =   990
         End
         Begin MSFlexGridLib.MSFlexGrid GridInfo 
            Height          =   1485
            Left            =   150
            TabIndex        =   35
            Top             =   225
            Width           =   6420
            _ExtentX        =   11324
            _ExtentY        =   2619
            _Version        =   393216
         End
         Begin VB.Frame Frame5 
            Caption         =   "Estoque"
            Height          =   750
            Left            =   165
            TabIndex        =   28
            Top             =   1800
            Width           =   6390
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Custo:"
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
               Left            =   2055
               TabIndex        =   38
               Top             =   345
               Width           =   555
            End
            Begin VB.Label LabelCustoEstoque 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2655
               TabIndex        =   37
               Top             =   315
               Width           =   975
            End
            Begin VB.Label LabelConsumoMedio 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   5265
               TabIndex        =   34
               Top             =   330
               Width           =   765
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Consumo Médio:"
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
               Left            =   3810
               TabIndex        =   33
               Top             =   390
               Width           =   1410
            End
            Begin VB.Label LabelSaldo 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1095
               TabIndex        =   32
               Top             =   315
               Width           =   795
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Saldo (KG):"
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
               Left            =   90
               TabIndex        =   30
               Top             =   360
               Width           =   990
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Importação"
            Height          =   1590
            Left            =   6690
            TabIndex        =   23
            Top             =   150
            Width           =   2670
            Begin VB.Label LabelParidade 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1635
               TabIndex        =   55
               Top             =   1125
               Width           =   900
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Paridade:"
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
               Left            =   735
               TabIndex        =   36
               Top             =   1200
               Width           =   825
            End
            Begin VB.Label LabelCustoCalculado 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1635
               TabIndex        =   27
               Top             =   690
               Width           =   900
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Custo Calculado:"
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
               TabIndex        =   26
               Top             =   750
               Width           =   1455
            End
            Begin VB.Label LabelCustoFOB 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1635
               TabIndex        =   25
               Top             =   225
               Width           =   900
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Custo FOB US$:"
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
               TabIndex        =   24
               Top             =   285
               Width           =   1395
            End
         End
      End
      Begin VB.Frame FrameCusto 
         Caption         =   "Custo"
         Height          =   1800
         Left            =   60
         TabIndex        =   13
         Top             =   870
         Width           =   9540
         Begin VB.CheckBox FreteNaoInf 
            Caption         =   "Obter da Última Entrada"
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
            Left            =   6690
            TabIndex        =   77
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2445
         End
         Begin VB.CheckBox CondPagtoNaoInf 
            Caption         =   "Obter da Última Compra"
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
            Left            =   6690
            TabIndex        =   76
            Top             =   810
            Value           =   1  'Checked
            Width           =   2445
         End
         Begin VB.CheckBox AliqNaoInf 
            Caption         =   "Obter da Última Compra"
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
            Left            =   6660
            TabIndex        =   75
            Top             =   240
            Value           =   1  'Checked
            Width           =   2445
         End
         Begin VB.Frame Frame8 
            Caption         =   "Valores Atuais"
            Height          =   1050
            Left            =   105
            TabIndex        =   70
            Top             =   615
            Width           =   3090
            Begin VB.Label LabelCustoAtual 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1605
               TabIndex        =   74
               Top             =   225
               Width           =   1275
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Custo:"
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
               Left            =   930
               TabIndex        =   73
               Top             =   300
               Width           =   555
            End
            Begin VB.Label LabelData 
               AutoSize        =   -1  'True
               Caption         =   "Atualizado Em:"
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
               TabIndex        =   72
               Top             =   720
               Width           =   1275
            End
            Begin VB.Label Data 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1605
               TabIndex        =   71
               Top             =   660
               Width           =   1275
            End
         End
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   4590
            TabIndex        =   14
            Top             =   810
            Width           =   1965
         End
         Begin MSMask.MaskEdBox AliquotaICMS 
            Height          =   315
            Left            =   4590
            TabIndex        =   15
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Custo 
            Height          =   315
            Left            =   1710
            TabIndex        =   16
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Frete 
            Height          =   315
            Left            =   4590
            TabIndex        =   17
            Top             =   1290
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCusto 
            AutoSize        =   -1  'True
            Caption         =   "Novo Custo:"
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
            Left            =   510
            TabIndex        =   21
            ToolTipText     =   "Para importação ignorar ICMS e Frete, nos outros casos este custo embute o ICMS mas não o frete."
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label LabelAliquotaICMS 
            AutoSize        =   -1  'True
            Caption         =   "Alíq.ICMS:"
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
            Left            =   3600
            TabIndex        =   20
            Top             =   285
            Width           =   930
         End
         Begin VB.Label LabelCondPagto 
            AutoSize        =   -1  'True
            Caption         =   "Cond.Pagto:"
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
            Left            =   3465
            TabIndex        =   19
            Top             =   900
            Width           =   1065
         End
         Begin VB.Label LabelFrete 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   4005
            TabIndex        =   18
            Top             =   1365
            Width           =   510
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Produto"
         Height          =   810
         Left            =   60
         TabIndex        =   7
         Top             =   -30
         Width           =   9525
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelUMEstoque 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8580
            TabIndex        =   12
            Top             =   285
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "U.M. Estoque:"
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
            Left            =   7260
            TabIndex        =   11
            Top             =   345
            Width           =   1230
         End
         Begin VB.Label LabelProduto 
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   10
            Top             =   345
            Width           =   735
         End
         Begin VB.Label Descricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2505
            TabIndex        =   9
            Top             =   300
            Width           =   4245
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   5835
      Index           =   2
      Left            =   240
      TabIndex        =   44
      Top             =   795
      Visible         =   0   'False
      Width           =   9690
      Begin VB.Frame Frame7 
         Caption         =   "Planilha"
         Height          =   4800
         Left            =   4920
         TabIndex        =   46
         Top             =   900
         Width           =   4710
         Begin VB.TextBox PlanImpValor 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   2580
            TabIndex        =   54
            Text            =   "Text1"
            Top             =   390
            Width           =   1410
         End
         Begin VB.TextBox PlanImpItem 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   165
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   375
            Width           =   2250
         End
         Begin MSFlexGridLib.MSFlexGrid GridPlanImp 
            Height          =   3825
            Left            =   120
            TabIndex        =   49
            Top             =   270
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   6747
            _Version        =   393216
         End
         Begin VB.Label LabelValorPlanImp 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1530
            TabIndex        =   60
            Top             =   4260
            Width           =   1395
         End
         Begin VB.Label LabelResulPlanImp 
            Caption         =   "Custo/KG R$:"
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
            Left            =   255
            TabIndex        =   59
            Top             =   4290
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Parâmetros"
         Height          =   4815
         Left            =   120
         TabIndex        =   45
         Top             =   885
         Width           =   4710
         Begin VB.CommandButton BotaoCalcular 
            Caption         =   "Calcular"
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
            Left            =   555
            TabIndex        =   57
            Top             =   3315
            Width           =   1605
         End
         Begin VB.CommandButton BotaoLimparGridParam 
            Caption         =   "Limpar"
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
            Left            =   2520
            TabIndex        =   58
            Top             =   3300
            Width           =   1605
         End
         Begin VB.TextBox ParamImpValor 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2535
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   390
            Width           =   1410
         End
         Begin VB.TextBox ParamImpItem 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   150
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   360
            Width           =   2250
         End
         Begin MSFlexGridLib.MSFlexGrid GridParamImp 
            Height          =   2820
            Left            =   135
            TabIndex        =   47
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4974
            _Version        =   393216
         End
         Begin VB.Label LabelDescrParam 
            BorderStyle     =   1  'Fixed Single
            Height          =   825
            Left            =   135
            TabIndex        =   48
            Top             =   3825
            Width           =   4455
         End
      End
      Begin VB.Label LabelProduto2 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   975
         TabIndex        =   63
         Top             =   345
         Width           =   1365
      End
      Begin VB.Label LabelDescricao2 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2400
         TabIndex        =   62
         Top             =   345
         Width           =   4245
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   61
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture 
      Height          =   555
      Left            =   7830
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CustoEmbMP.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CustoEmbMP.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CustoEmbMP.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CustoEmbMP.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   6630
      Left            =   120
      TabIndex        =   5
      Top             =   435
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   11695
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cadastro"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "planilha de importação"
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
Attribute VB_Name = "CustoEmbMPOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private dAliqAnt As Double
Private dFreteAnt As Double
Private dCustoMoedaAnt As Double
Private dCotacaoAnt As Double
Private iMoedaAnt As Integer

Dim gcolPlanilhasImp As Collection
Dim gcolParamPlanImp As Collection

Dim m_Caption As String
Event Unload()

'evento do browse de produto
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCusto As AdmEvento
Attribute objEventoCusto.VB_VarHelpID = -1

Dim objGridInfo As AdmGrid
Dim iGrid_CustoInfo As Integer
Dim iGrid_DataInfo As Integer
Dim iGrid_MoedaInfo As Integer
Dim iGrid_FreteInfo As Integer
Dim iGrid_ICMSInfo As Integer
Dim iGrid_TaxaInfo As Integer

Dim objGridParamImp As AdmGrid
Dim iGrid_ParamImpItem As Integer
Dim iGrid_ParamImpValor As Integer

Dim objGridPlanImp As AdmGrid
Dim iGrid_PlanImpItem As Integer
Dim iGrid_PlanImpValor As Integer

'variaveis de controle de alteração
Public iAlterado As Integer
Dim iProdutoAlterado As Integer

Dim iFrameAtual As Integer

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCustoEmb As New ClassCustoEmbMP

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CustoEmbMP"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objCustoEmb)
    If lErro <> SUCESSO Then gError 116286
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objCustoEmb.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "DataAtualizacao", objCustoEmb.dtDataAtualizacao, 0, "DataAtualizacao"
    colCampoValor.Add "Custo", objCustoEmb.dCusto, 0, "Custo"
    colCampoValor.Add "FretePorKg", objCustoEmb.dFretePorKg, 0, "FretePorKg"
    colCampoValor.Add "AliquotaICMS", objCustoEmb.dAliquotaICMS, 0, "AliquotaICMS"
    colCampoValor.Add "CondicaoPagto", objCustoEmb.iCondicaoPagto, 0, "CondicaoPagto"
    
    'adiciona FilialEmpresa
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub
    
Erro_Tela_Extrai:
    
    Select Case gErr

        Case 116286

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158587)
            
    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(ByVal objCustoEmb As ClassCustoEmbMP) As Long
'Move os dados da tela p/ a memoria

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria
    
    'verifica se o produto esta preenchido
    If Len(Trim(Produto.ClipText)) > 0 And Len(Trim(Custo.Text)) > 0 Then
         
        'Retira a mascara do produto
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116287
    
        objCustoEmb.sProduto = sProduto
        
        'verifica se a aliquota esta em branco
        If Len(Trim(AliquotaICMS.Text)) = 0 Then
            objCustoEmb.dAliquotaICMS = StrParaDbl(AliquotaICMS.Text)
        Else
            objCustoEmb.dAliquotaICMS = StrParaDbl(AliquotaICMS.Text / 100)
        End If
        
        objCustoEmb.dCusto = StrParaDbl(Custo.Text)
        objCustoEmb.dFretePorKg = StrParaDbl(Frete.Text)
        objCustoEmb.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
        objCustoEmb.iFilialEmpresa = giFilialEmpresa
        objCustoEmb.dtDataAtualizacao = gdtDataHoje
        objCustoEmb.iAliquotaICMSInf = IIf(AliqNaoInf.Value = vbChecked, 0, 1)
        objCustoEmb.iCondicaoPagtoInf = IIf(CondPagtoNaoInf.Value = vbChecked, 0, 1)
        objCustoEmb.iFretePorKGInf = IIf(FreteNaoInf.Value = vbChecked, 0, 1)
    
        '####################################
        'Inserido por Wagner 07/12/05
        lErro = Move_ParamImp_Memoria(objCustoEmb)
        If lErro <> SUCESSO Then gError 141324
        '####################################
    
    End If
        
    Move_Tela_Memoria = SUCESSO
        
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 116287, 141324 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158588)
    
    End Select
    
    Exit Function
    
End Function

Public Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCustoEmb As New ClassCustoEmbMP
Dim objProduto As New ClassProduto

On Error GoTo Erro_Tela_Preenche

    'preenche o obj c/ os valores correspondentes
    objCustoEmb.sProduto = colCampoValor.Item("Produto").vValor
    objCustoEmb.dtDataAtualizacao = colCampoValor.Item("DataAtualizacao").vValor
    objCustoEmb.dCusto = colCampoValor.Item("Custo").vValor
    objCustoEmb.dFretePorKg = colCampoValor.Item("FretePorKg").vValor
    objCustoEmb.dAliquotaICMS = colCampoValor.Item("AliquotaICMS").vValor
    objCustoEmb.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
    
    'Guarda o código do produto em objproduto
    objProduto.sCodigo = objCustoEmb.sProduto
    
    'Critica o formato do codigo
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 116297
    
    'Se não encontrou o produto => erro
    If lErro = 28030 Then gError 116341
            
    'Traz os dados para tela
    lErro = Traz_CustoEmb_Tela2(objCustoEmb, objProduto)
    If lErro <> SUCESSO Then gError 116288

    Exit Function

Erro_Tela_Preenche:

    Select Case gErr

        Case 116288, 116341

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158589)

    End Select

    Exit Function

End Function

Private Function Traz_CustoEmb_Tela() As Long
'rotina que traz os dados da tela

Dim lErro As Long
Dim objCustoEmb As New ClassCustoEmbMP
Dim objProduto As New ClassProduto

On Error GoTo Erro_Traz_CustosEmb_Tela
    
    'Critica o produto
    lErro = Traz_CustoEmb_Tela1(objCustoEmb, objProduto)
    If lErro <> SUCESSO Then gError 116342
    
    'lê os dados em CustoEmbMP e Exibe os dados na tela
    lErro = Traz_CustoEmb_Tela2(objCustoEmb, objProduto)
    If lErro <> SUCESSO Then gError 116343
    
    Traz_CustoEmb_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_CustosEmb_Tela:

    Traz_CustoEmb_Tela = gErr

    Select Case gErr
        
        Case 116342, 116343, 141323
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158590)
    
    End Select

    Exit Function

End Function

Private Function Traz_CustoEmb_Tela1(ByVal objCustoEmb As ClassCustoEmbMP, ByVal objProduto As ClassProduto) As Long
'Critica o produto

Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_Traz_CustoEmb_Tela1

    'Critica o formato do codigo
    lErro = CF("Produto_Critica_Filial", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 116297
            
    'lErro = 25041 => inexistente
    If lErro = 51381 Then gError 116298
    
    Traz_CustoEmb_Tela1 = SUCESSO
    
    Exit Function
    
Erro_Traz_CustoEmb_Tela1:

    Traz_CustoEmb_Tela1 = gErr
    
    Select Case gErr
        
        Case 116297
         
        Case 116298
           'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                Call Limpa_Tela_CustoEmbMP(True)
                Produto.SetFocus
            End If
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158591)
            
    End Select

End Function

Private Function Traz_CustoEmb_Tela2(ByVal objCustoEmb As ClassCustoEmbMP, ByVal objProduto As ClassProduto) As Long
'Exibe os dados na tela

Dim lErro As Long, dtDataAux As Date, dFreteAux As Double

On Error GoTo Erro_Traz_CustosEmb_Tela2
    
    'limpa tela, exceto o produto
    Call Limpa_Tela_CustoEmbMP(False)
            
    '************Lê os dados em CustoEmbMP*****************
    'preenche o obj com os dados a serem lidos
    objCustoEmb.sProduto = objProduto.sCodigo
    objCustoEmb.iFilialEmpresa = giFilialEmpresa
    
    'verifica se o produto tem alguma relação c/ custo
    lErro = CF("CustoEmbMP_Le", objCustoEmb)
    If lErro <> SUCESSO And lErro <> 116309 Then gError 116299
    '**********************************************************
    
    '************Preenche o frame Produto (labels)******************
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
    
    Descricao.Caption = objProduto.sDescricao
    
    LabelProduto2.Caption = Produto.Text
    
    LabelDescricao2.Caption = Descricao.Caption
    
    LabelUMEstoque.Caption = objProduto.sSiglaUMEstoque
    '******************************************************
            
    '******************Preenche o restante da tela*************
    'se não tiver, não preenche a data
    If objCustoEmb.dtDataAtualizacao = 0 Then
        Data.Caption = ""
        LabelCustoAtual = ""
    Else
        'preenche a label data c/ a data de atualização
        Data.Caption = Format(objCustoEmb.dtDataAtualizacao, "dd/mm/yy")
        LabelCustoAtual.Caption = Format(objCustoEmb.dCusto, FORMATO_CUSTO)
    End If
    
    'preenche o frete
    If objCustoEmb.iFretePorKGInf <> 0 Then
        Frete.Text = objCustoEmb.dFretePorKg
        FreteNaoInf = vbUnchecked
    Else
        FreteNaoInf = vbChecked
    End If
    
    'preenche a aliquotaICMS
    If objCustoEmb.iAliquotaICMSInf <> 0 Then
        AliquotaICMS.Text = (objCustoEmb.dAliquotaICMS * 100)
        AliqNaoInf = vbUnchecked
    Else
        AliqNaoInf = vbChecked
    End If
    
    'prenche a condição de pagto
    If objCustoEmb.iCondicaoPagtoInf <> 0 Then
        If objCustoEmb.iCondicaoPagto = 0 Then
            CondicaoPagamento.Text = ""
        Else
            CondicaoPagamento.Text = objCustoEmb.iCondicaoPagto
            Call CondicaoPagamento_Validate(bSGECancelDummy)
        End If
        CondPagtoNaoInf = vbUnchecked
    Else
        CondPagtoNaoInf = vbChecked
    End If
    
    lErro = Traz_CustoEmb_Tela3(objProduto)
    If lErro <> SUCESSO Then gError 106711
    
    'preenche a labelMoeda com o tipo de moeda p/ compra de produto
    lErro = ObterMoeda(objCustoEmb.sProduto)
    If lErro <> SUCESSO Then gError 123013
    
    '####################################
    'Inserido por Wagner 07/12/05
    lErro = Preenche_ParamImp(objCustoEmb)
    If lErro <> SUCESSO Then gError 141323
    '####################################
    
    'preenche o custo
    Call Acha_Maior
    
    iAlterado = 0
    iProdutoAlterado = 0
    
    Traz_CustoEmb_Tela2 = SUCESSO
    
    Exit Function
    
Erro_Traz_CustosEmb_Tela2:

    Traz_CustoEmb_Tela2 = gErr

    Select Case gErr
        
        Case 106711, 123013, 116299, 141323
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158592)
    
    End Select

    Exit Function

End Function

Sub Acha_Maior()
'Preenche o campo custo com o maior custo

Dim lErro As Long, objProduto As New ClassProduto
Dim dCustoEstoque As Double, dCustoUltCot As Double, dCustoUltCom As Double, dCustoImp As Double, dCustoUltEnt As Double, dCustoNacMoeda As Double
Dim sProduto As String, dTaxa As Double
Dim iProdutoPreenchido As Integer
Dim objFilial As AdmFiliais

On Error GoTo Erro_Acha_Maior

    'verifica se o produto esta preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then
         
        'Retira a mascara do produto
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 106908
    
        objProduto.sCodigo = sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 106907
        If lErro = SUCESSO Then
    
            'Verifica qual dos custos é o maior e preenche o campo custo
            
            dCustoEstoque = StrParaDbl(LabelCustoEstoque.Caption)
            'se nao for mercadoria importada
            If objProduto.iOrigemMercadoria <> 1 Then
                
                dCustoEstoque = dCustoEstoque - StrParaDbl(Frete.Text)
                
                Set objFilial = New AdmFiliais
                
                'fazer a leitura
                objFilial.iCodFilial = giFilialEmpresa
                lErro = CF("FilialEmpresa_Le", objFilial)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                If objFilial.iSuperSimples = 0 Then
                    dCustoEstoque = dCustoEstoque / (1 - StrParaDbl(AliquotaICMS.Text) / 100)
                End If
                
            End If
            
            dTaxa = StrParaDbl(Cotacao.Text)
            dCustoUltCot = StrParaDbl(GridInfo.TextMatrix(1, iGrid_CustoInfo)) * IIf(GridInfo.TextMatrix(1, iGrid_MoedaInfo) <> "", dTaxa, 1)
            dCustoUltCom = StrParaDbl(GridInfo.TextMatrix(2, iGrid_CustoInfo)) * IIf(GridInfo.TextMatrix(2, iGrid_MoedaInfo) <> "", dTaxa, 1)
            dCustoImp = StrParaDbl(LabelCustoCalculado.Caption)
            dCustoUltEnt = StrParaDbl(GridInfo.TextMatrix(3, iGrid_CustoInfo))
            dCustoNacMoeda = StrParaDbl(LabelCustoCalculado2.Caption)
            
            If dCustoEstoque >= dCustoImp And dCustoEstoque >= dCustoUltCot And dCustoEstoque >= dCustoUltEnt And dCustoEstoque >= dCustoUltCot And dCustoEstoque >= dCustoNacMoeda Then
                
                Custo.Text = Format(dCustoEstoque, FORMATO_CUSTO)
                
            ElseIf dCustoImp >= dCustoUltCot And dCustoImp >= dCustoUltEnt And dCustoImp >= dCustoUltCot And dCustoImp >= dCustoNacMoeda Then
            
                Custo.Text = Format(dCustoImp, FORMATO_CUSTO)
                    
            ElseIf dCustoUltCot >= dCustoUltEnt And dCustoUltCot >= dCustoUltCot And dCustoUltCot >= dCustoNacMoeda Then
            
                Custo.Text = Format(dCustoUltCot, FORMATO_CUSTO)
                
            ElseIf dCustoUltEnt >= dCustoUltCot And dCustoUltEnt >= dCustoNacMoeda Then
            
                Custo.Text = Format(dCustoUltEnt, FORMATO_CUSTO)
                
            ElseIf dCustoUltCot >= dCustoNacMoeda Then
            
                Custo.Text = Format(dCustoUltCom, FORMATO_CUSTO)
            
            Else
            
                Custo.Text = Format(dCustoNacMoeda, FORMATO_CUSTO)
                
            End If
    
        End If
            
    End If
    
    Exit Sub
     
Erro_Acha_Maior:

    Select Case gErr
          
        Case 106907, 106908, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158593)
     
    End Select
     
    Exit Sub
    
End Sub

Function ObterMoeda(sProduto As String) As Long
'Preenche a label Moeda com a moeda que sera utilizada  para a compra do material

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim iValor As Integer, sNome As String

On Error GoTo Erro_ObterMoeda

    sNome = String(STRING_NOME_MOEDA, 0)
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = SUCESSO Then gError 123009
    
    If Moeda.ListIndex = -1 Then
    
        'Seleciona o campo valor1 da tabela CategoriaProdutoItem
        lErro = Comando_Executar(lComando, "SELECT Valor1, Moedas.Nome FROM Moedas, ProdutoCategoria, CategoriaProdutoItem WHERE Valor2 = 1 AND Moedas.Codigo = Valor1 AND ProdutoCategoria.Categoria = CategoriaProdutoItem.Categoria AND ProdutoCategoria.Item = CategoriaProdutoItem.Item AND Produto = ? AND ProdutoCategoria.Categoria=?", iValor, sNome, sProduto, gobjFAT.sCategMoeda)
        If lErro <> AD_SQL_SUCESSO Then gError 123010
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 123011
        
        If lErro = AD_SQL_SUCESSO And iValor <> 0 Then
        
            LabelMoeda.Caption = sNome
            
            lErro = ObterCotacao(iValor)
            If lErro <> SUCESSO Then gError 123014
        
        Else
            
            LabelMoeda.Caption = "Dólar"
            lErro = ObterCotacao(MOEDA_DOLAR)
            If lErro <> SUCESSO Then gError 123014
            
            iValor = MOEDA_DOLAR
            
        End If
        
        Call Combo_Seleciona_ItemData(Moeda, iValor)
    
    Else
        lErro = ObterCotacao(Codigo_Extrai(Moeda))
        If lErro <> SUCESSO Then gError 123014
    End If
    
    'realiza o fechamento do comando
    Call Comando_Fechar(lComando)
    
    'retorna sucesso
    ObterMoeda = SUCESSO
    
    Exit Function
    
Erro_ObterMoeda:
    
    ObterMoeda = gErr

    Select Case gErr
        
        Case 123009
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 123010, 123011
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCATEGORIA", gErr, sProduto)
        
        Case 123014
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158594)

    End Select
    
    'realiza o fechamento do comando
    Call Comando_Fechar(lComando)

    Exit Function
   
End Function

Function ObterCotacao(iValor As Integer) As Long
'Preenche a TextBox cotacao

Dim lErro As Long
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda
Dim NewDate As Date, objMnemonicoFPTipo As New ClassMnemonicoFPTipo

On Error GoTo Erro_ObterCotacao

    If iValor = MOEDA_DOLAR Then
    
        'obter valor do mnemonico TaxaDolar
        With objMnemonicoFPTipo
            .iTipoPlanilha = PLANILHA_TIPO_TODOS
            .iFilialEmpresa = giFilialEmpresa
            .iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
            .sItemCategoria = ""
            .sProduto = ""
            .iTabelaPreco = 0
            .sMnemonico = "TaxaDolar"
        End With
        lErro = CF("MnemonicoFPTipo_Le", objMnemonicoFPTipo)
        If lErro <> SUCESSO And lErro <> 106912 Then gError 106913
        If lErro = SUCESSO Then
            Cotacao.Text = objMnemonicoFPTipo.sExpressao
        End If
        
    ElseIf iValor = MOEDA_REAL Then
        Cotacao.Text = Format(1, "###,##0.00###")
    Else
    
        'Carrega objCotacao
        objCotacao.dtData = gdtDataAtual
        objCotacao.iMoeda = iValor
        
        'Chama função de leitura
        lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
        If lErro <> SUCESSO Then gError 123012
        
        'Verifica se o objCotacao está preenchido
        If objCotacao.dValor = 0 Then
            Cotacao.Text = ""
        Else
            Cotacao.Text = Format(objCotacao.dValor, "###,##0.00###")
        End If
    
    End If
    
    Call Cotacao_Validate(bSGECancelDummy)
    
    ObterCotacao = SUCESSO
        
    Exit Function
    
Erro_ObterCotacao:

    ObterCotacao = gErr
    
    Select Case gErr
    
        Case 123012, 106913
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158595)

    End Select
    
    Exit Function
        
End Function

Private Function Carrega_CondicaoPagamento() As Long
'Carrega as Condições de Pagamento

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 116289

    For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo ítem na List da Combo CondicaoPagamento
        CondicaoPagamento.AddItem StrParaInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case 116289

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158596)

    End Select

    Exit Function

End Function

Sub Preenche_CustoCalculado2()
'Preenche o custo calculado do Material comprado em mercado nacional

Dim dValor As Double

    'Verifica se o custoMoeda foi preenchido e a Cotacao tambem
    If Len(Trim(CustoMoeda.Text)) <> 0 And Len(Trim(Cotacao.Text)) <> 0 Then
    
        dValor = StrParaDbl(CustoMoeda.Text) * StrParaDbl(Cotacao.Text)
        
        LabelCustoCalculado2.Caption = Format(dValor, "###,##0.00###")
        
    Else
    
        LabelCustoCalculado2.Caption = ""
        
    End If
    
    'Chama a funcao que ira verificar o maior custo
    Call Acha_Maior
        
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 116290
    
    'inicializa o evento de browser
    Set objEventoProduto = New AdmEvento
    Set objEventoCusto = New AdmEvento
    
    'Carrega as Condições de Pagamento
    lErro = Carrega_CondicaoPagamento()
    If lErro <> SUCESSO Then gError 116291

    Set objGridInfo = New AdmGrid
    lErro = Inicializa_Grid_Info(objGridInfo)
    If lErro <> SUCESSO Then gError 106634
    
    Set objGridParamImp = New AdmGrid
    lErro = Inicializa_Grid_ParamImp(objGridParamImp)
    If lErro <> SUCESSO Then gError 106635
    
    lErro = ParamImp_CarregaGrid
    If lErro <> SUCESSO Then gError 106706
    
    Set objGridPlanImp = New AdmGrid
    lErro = Inicializa_Grid_PlanImp(objGridPlanImp)
    If lErro <> SUCESSO Then gError 106636
    
    lErro = PlanImp_CarregaGrid
    If lErro <> SUCESSO Then gError 106707
    
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
            
        Case 116290, 116291, 106634, 106635, 106636, 106706, 106707
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158597)

    End Select
    
    Exit Sub
    
End Sub

Private Sub AliquotaICMS_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(AliquotaICMS)
    dAliqAnt = StrParaDbl(AliquotaICMS.Text)

End Sub

Private Sub AliquotaICMS_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AliquotaICMS_Validate(Cancel As Boolean)
'verifica se o valor do ICMS é valido

Dim lErro As Long

On Error GoTo Erro_Custo_Validate

    If Len(Trim(AliquotaICMS.Text)) <> 0 Then

        'testa para ver se é uma porcentagem valida
        lErro = Porcentagem_Critica(AliquotaICMS.Text)
        If lErro <> SUCESSO Then gError 116292

        'AliquotaICMS.Text = Format(AliquotaICMS.Text, "Fixed")

    End If
    
    If Abs(StrParaDbl(AliquotaICMS.Text) - dAliqAnt) > DELTA_VALORMONETARIO Then Call Acha_Maior
    
    Exit Sub

Erro_Custo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116292

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158598)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoCalcular_Click()

Dim iLinha As Integer, objPlanilhas As ClassPlanilhas
Dim lErro As Long, dValor As Double
Dim colMnemonicoValor As New ClassColMnemonicoValor
Dim objMnemonicoValor As ClassMnemonicoValor
Dim objContexto As New ClassContextoPlan, iProdutoPreenchido As Integer, sProduto As String

On Error GoTo Erro_BotaoCalcular_Click

    For iLinha = 1 To gcolParamPlanImp.Count
    
        If GridParamImp.TextMatrix(iLinha, iGrid_ParamImpValor) = "" Then gError 106708
        
        Set objMnemonicoValor = New ClassMnemonicoValor
        Set objMnemonicoValor.colValor = New Collection
        
        objMnemonicoValor.sMnemonico = GridParamImp.TextMatrix(iLinha, iGrid_ParamImpItem)
        objMnemonicoValor.colValor.Add CDbl(GridParamImp.TextMatrix(iLinha, iGrid_ParamImpValor))
        
        objContexto.colMnemonicoValor.Add objMnemonicoValor
        
    Next
    
    objContexto.iFilialFaturamento = giFilialEmpresa
    
    'verifica se o produto esta preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then
         
        'Retira a mascara do produto
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 106710
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objContexto.sProduto = sProduto
    
    End If
    
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco3", gcolPlanilhasImp, dValor, objContexto)
    If lErro <> SUCESSO Then gError 106709

    For Each objPlanilhas In gcolPlanilhasImp
    
        If objPlanilhas.sExpressao = "FOB" Then LabelCustoFOB.Caption = Format(objPlanilhas.dValor, "###,##0.00###")
        If objPlanilhas.sExpressao = "TaxaDolar" Then LabelParidade.Caption = Format(objPlanilhas.dValor, "###,##0.00###")
        
        GridPlanImp.TextMatrix(objPlanilhas.iLinha, iGrid_PlanImpValor) = Format(objPlanilhas.dValor, "###,##0.00###")

    Next
        
    LabelValorPlanImp.Caption = Format(dValor, "###,##0.00###")
    LabelCustoCalculado.Caption = Format(dValor, "###,##0.00###")
    
    'Chama a funcao que ira preencher o maior custo
    Call Acha_Maior
    
    Exit Sub
    
Erro_BotaoCalcular_Click:

    Select Case gErr
          
        Case 106708
            Call Rotina_Erro(vbOKOnly, "ERRO_PARAMETRO_NAO_PREENCHIDO_LINHA", gErr, iLinha)
        
        Case 106709
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158599)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long
        
On Error GoTo Erro_Botao_Fechar
        
    'pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 116340
    
    Unload Me
    
    Exit Sub
    
Erro_Botao_Fechar:

    Select Case gErr
    
        Case 116340
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158600)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
'sub para limpar a tela

Dim lErro As Long

On Error GoTo Erro_Botao_Limpar

    'pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 116293
    
    'limpa a tela
    Call Limpa_Tela_CustoEmbMP(True)
    
    Exit Sub
        
Erro_Botao_Limpar:

    Select Case gErr

        Case 116293
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158601)

    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_CustoEmbMP(Optional bLimpaTudo As Boolean = True)
'sub que limpa a tela inteira ou apenas o frame custo

Dim sProduto As String, iLinha As Integer

On Error GoTo Erro_Limpa_Tela_CustoEmbMP

    sProduto = Produto.Text
    
    'limpa as text box
    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridInfo)
    
    LabelCustoEstoque.Caption = ""
    LabelSaldo.Caption = ""
    LabelConsumoMedio.Caption = ""
    LabelCustoFOB.Caption = ""
    LabelParidade.Caption = ""
    LabelValorPlanImp.Caption = ""
    LabelCustoCalculado.Caption = ""
    LabelCustoCalculado2.Caption = ""
    
    'limpa o restante
    Descricao.Caption = ""
    LabelUMEstoque.Caption = ""
    Data.Caption = ""
    LabelCustoAtual = ""
    CondicaoPagamento.Text = ""
    LabelProduto2.Caption = ""
    LabelDescricao2.Caption = ""
    LabelMoeda.Caption = ""
    LabelDescrParam.Caption = ""
    
    'limpar valores de parametros de importacao
    For iLinha = 1 To objGridParamImp.iLinhasExistentes
    
        GridParamImp.TextMatrix(iLinha, iGrid_ParamImpValor) = ""
        
    Next
    
    'limpar valores da planilha de importacao
    For iLinha = 1 To objGridPlanImp.iLinhasExistentes
    
        GridPlanImp.TextMatrix(iLinha, iGrid_PlanImpValor) = ""
    
    Next
    
    iAlterado = 0
    iProdutoAlterado = 0
    
    If bLimpaTudo = False Then
        Produto.Text = sProduto
        Exit Sub
    End If

    Exit Sub
    
Erro_Limpa_Tela_CustoEmbMP:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158602)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimparGridParam_Click()

Dim iLinha As Integer

    For iLinha = 1 To gcolParamPlanImp.Count
    
        GridParamImp.TextMatrix(iLinha, iGrid_ParamImpValor) = ""
        
    Next
    
    For iLinha = 1 To gcolPlanilhasImp.Count
    
        GridPlanImp.TextMatrix(iLinha, iGrid_PlanImpValor) = ""
    
    Next
    
End Sub

Private Sub BotaoConsultaCusto_Click()

Dim objCusto As ClassCustoEmbMP
Dim colSelecao As Collection

    'chama a tela de custo
    Call Chama_Tela("CustoEmbMPLista", colSelecao, objCusto, objEventoCusto)

End Sub

Private Sub Cotacao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Cotacao_GotFocus()

    Call MaskEdBox_TrataGotFocus(Cotacao)
    dCotacaoAnt = StrParaDbl(Cotacao.Text)
    
End Sub

Private Sub Cotacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cotacao_Validate

    'Verifica se foi preenchido a Cotacao
    If Len(Trim(Cotacao.Text)) <> 0 Then

        'não pode ser nº negativo
        lErro = Valor_NaoNegativo_Critica(Cotacao.Text)
        If lErro <> SUCESSO Then gError 123009

    End If

    If Abs(dCotacaoAnt - StrParaDbl(Cotacao.Text)) > DELTA_VALORMONETARIO Then Call Preenche_CustoCalculado2
        
    Exit Sub

Erro_Cotacao_Validate:

    Cancel = True
    
    Select Case gErr

        Case 123009

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158603)
    
    End Select

    Exit Sub
    
End Sub

Private Sub Custo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Custo)

End Sub

Private Sub Custo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Custo_Validate(Cancel As Boolean)
'verifica se o valor de Custo é valido

Dim lErro As Long

On Error GoTo Erro_Custo_Validate

    If Len(Trim(Custo.Text)) <> 0 Then

        'não pode ser nº negativo
        lErro = Valor_NaoNegativo_Critica(Custo.Text)
        If lErro <> SUCESSO Then gError 116294

    End If

    Exit Sub

Erro_Custo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116294

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158604)
    
    End Select

    Exit Sub

End Sub

Private Sub CustoMoeda_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CustoMoeda_GotFocus()

    Call MaskEdBox_TrataGotFocus(CustoMoeda)
    dCustoMoedaAnt = StrParaDbl(CustoMoeda.Text)
    
End Sub

Private Sub CustoMoeda_Validate(Cancel As Boolean)
'verifica se o valor de CustoMoeda é valido

Dim lErro As Long

On Error GoTo Erro_CustoMoeda_Validate

    'Verifica se foi preenchido o CustoMoeda
    If Len(Trim(CustoMoeda.Text)) <> 0 Then

        'não pode ser nº negativo
        lErro = Valor_NaoNegativo_Critica(CustoMoeda.Text)
        If lErro <> SUCESSO Then gError 123007

    End If

    If Abs(dCustoMoedaAnt - StrParaDbl(CustoMoeda.Text)) > DELTA_VALORMONETARIO Then Call Preenche_CustoCalculado2
        
    Exit Sub

Erro_CustoMoeda_Validate:

    Cancel = True
    
    Select Case gErr

        Case 123007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158605)
    
    End Select

    Exit Sub
    
End Sub

Private Sub Frete_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Frete)
    dFreteAnt = StrParaDbl(Frete.Text)
    
End Sub

Private Sub Frete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Frete_Validate(Cancel As Boolean)
'verifica se o valor do frete é valido

Dim lErro As Long

On Error GoTo Erro_Frete_Validate

    If Len(Trim(Frete.Text)) <> 0 Then

        'não pode ser nº negativo
        lErro = Valor_NaoNegativo_Critica(Frete.Text)
        If lErro <> SUCESSO Then gError 116295

    End If

    If Abs(StrParaDbl(Frete.Text) - dFreteAnt) > DELTA_VALORMONETARIO Then Call Acha_Maior
    
    Exit Sub

Erro_Frete_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116295

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158606)
    
    End Select

    Exit Sub

End Sub

Private Sub LabelProduto_Click()
'sub chamadora do browser Produto

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'formata o produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116296

        'Preenche o código de objProduto
        objProduto.sCodigo = sProdutoFormatado

    End If

    'chama a tela de produtos
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 116296

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158607)

    End Select

    Exit Sub

End Sub

Private Sub Produto_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Produto)

End Sub

Private Sub Produto_Change()
           
    iAlterado = REGISTRO_ALTERADO
    iProdutoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Validate(Cancel As Boolean)
'Valida o código do produto e preenche o restante da tela

Dim lErro As Long

On Error GoTo Erro_Produto_Validate

    'se nao houve alteracao na produto, sai da rotina
    If iProdutoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'se o codigo estiver vazio  => sai da rotina
    If Len(Trim(Produto.ClipText)) = 0 Then
        Call Limpa_Tela_CustoEmbMP(True)
        Exit Sub
    End If
    
    lErro = Traz_CustoEmb_Tela()
    If lErro <> SUCESSO Then gError 116339
        
    Exit Sub

Erro_Produto_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 116339
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158608)
            
    End Select
    
    Exit Sub

End Sub

Private Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagamento_Click()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_CondicaoPagamento_Click
    
    'Verifica se alguma Condição foi selecionada
    If CondicaoPagamento.ListIndex = -1 Then Exit Sub

    'Passa o código da Condição para objCondicaoPagto
    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)

    'Lê Condição a partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 116300
    
    'Se não achou a Condição de Pagamento --> erro
    If lErro = 19205 Then gError 116301

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 116300

        Case 116301
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158609)

      End Select

    Exit Sub

End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)
'verifica se a condição de pagto é valida

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Condicaopagamento_Validate

    'Verifica se a Condicao Pagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicao Pagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 116302

    'Se não encontra, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 116303
        If lErro <> SUCESSO Then gError 116304

        'Coloca na tela
        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida
        
    End If

    'Não encontrou e é STRING
    If lErro = 6731 Then gError 116305

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True
    
    Select Case gErr

       Case 116302, 116303

       Case 116304
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 116305
             Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158610)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Sub para inicializar a rotina de gravação

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'grava o registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 116310

    'limpa a tela(toda)
    Call Limpa_Tela_CustoEmbMP(True)
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116310

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158611)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Função para gravar registro no BD

Dim objCustoEmb As New ClassCustoEmbMP
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    'tranforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
     
    'verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 116312
    
    'verifica se o custo foi preenchido
    If StrParaDbl(Custo.Text) = 0 Then gError 116335
    
    'move os dados da tela p/ o obj
    lErro = Move_Tela_Memoria(objCustoEmb)
    If lErro <> SUCESSO Then gError 116313
            
    'grava o registro no Bd
    lErro = CF("CustoEmbMP_Grava", objCustoEmb)
    If lErro <> SUCESSO Then gError 116311
    
    'volta o ponteiro ao padrao
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:
 
    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
 
    Select Case gErr
     
        Case 116313, 116311
     
        Case 116312
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus
                
        Case 116335
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTO_NAO_PREENCHIDO", gErr)
            Custo.SetFocus
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158612)
    
    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Sub que inicializa a exclusão de registros

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCustoEmb As New ClassCustoEmbMP
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoExcluir_Click
    
    'Verifica preenchimento do codigo
    If Len(Trim(Produto.ClipText)) = 0 Then gError 116321

    'Retira a mascara do produto
    lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 116337

    'preenche o obj c/ os dados a serem passados como parametro
    objCustoEmb.sProduto = sProduto
    objCustoEmb.iFilialEmpresa = giFilialEmpresa
    
    'LE a tabela CustoEmbMP a relação custo/produto
    lErro = CF("CustoEmbMP_Le", objCustoEmb)
    If lErro <> SUCESSO And lErro <> 116309 Then gError 116322

    'Se não achou --> Erro
    If lErro = 116309 Then gError 116323
    
    'pergunta se relamente deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CUSTOEMBMP", objCustoEmb.sProduto)

    'se sim
    If vbMsgRes = vbYes Then
        
        'tranforma o ponteiro em ampulheta
        GL_objMDIForm.MousePointer = vbHourglass
        
        'exclui o registro
        lErro = CF("CustoEmbMP_Exclui", objCustoEmb)
        If lErro <> SUCESSO Then gError 116324
                                                        
        'limpa a tela(toda)
        Call Limpa_Tela_CustoEmbMP(True)
            
        'volta o ponteiro ao padrão
        GL_objMDIForm.MousePointer = vbDefault
    
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 116322, 116324, 116337

        Case 116321
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus
        
        Case 116323
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTO_NAO_EXISTENTE", gErr, Produto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158613)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser Produto

Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    'Preenche campo Produto
    Produto.PromptInclude = False
    Produto.Text = CStr(objProduto.sCodigo)
    Produto.PromptInclude = True
    Produto_Validate (bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158614)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoCusto_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser Custo

Dim objCustoEmb As ClassCustoEmbMP
Dim lErro As Long

On Error GoTo Erro_objEventoCusto_evSelecao

    Set objCustoEmb = obj1
    
    'Exibe o código do produto na tela
    Produto.PromptInclude = False
    Produto.Text = objCustoEmb.sProduto
    Produto.PromptInclude = True
        
    'traz as informações do BD para a tela
    lErro = Traz_CustoEmb_Tela()
    If lErro <> SUCESSO Then gError 123006


    Me.Show

    Exit Sub

Erro_objEventoCusto_evSelecao:

    Select Case gErr

        Case 123006
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158615)

    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objCustoEmb As ClassCustoEmbMP) As Long
'Espera receber o código do produto em objCustoEmb

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum parametro
    If Not (objCustoEmb Is Nothing) Then
    
        'Exibe o código do produto na tela
        Produto.PromptInclude = False
        Produto.Text = objCustoEmb.sProduto
        Produto.PromptInclude = True
        
        'traz as informações do BD para a tela
        lErro = Traz_CustoEmb_Tela()
        If lErro <> SUCESSO Then gError 116334

    End If

    iAlterado = 0
    iProdutoAlterado = 0
        
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
        
        Case 116334
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158616)
        
    End Select
        
    Exit Function
        
End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        End If
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set gcolPlanilhasImp = Nothing
    Set gcolParamPlanImp = Nothing
    Set objGridInfo = Nothing
    Set objGridParamImp = Nothing
    Set objGridPlanImp = Nothing
    
    Set objEventoProduto = Nothing
    Set objEventoCusto = Nothing
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Custos de Matérias Primas e Embalagens"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CustoEmbMP"
    
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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Produto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Produto, Button, Shift, X, Y)
End Sub

Private Sub LabelAliquotaICMS_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAliquotaICMS, Source, X, Y)
End Sub

Private Sub LabelAliquotaICMS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAliquotaICMS, Button, Shift, X, Y)
End Sub

Private Sub LabelCondPagto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCondPagto, Source, X, Y)
End Sub

Private Sub LabelCondPagto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCondPagto, Button, Shift, X, Y)
End Sub

Private Sub LabelCusto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCusto, Source, X, Y)
End Sub

Private Sub LabelCusto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCusto, Button, Shift, X, Y)
End Sub

Private Sub LabelData_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub

Private Sub LabelData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFrete, Source, X, Y)
End Sub

Private Sub LabelFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFrete, Button, Shift, X, Y)
End Sub

Private Sub LabelUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelUMEstoque, Source, X, Y)
End Sub

Private Sub LabelUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelUMEstoque, Button, Shift, X, Y)
End Sub

Private Function Inicializa_Grid_Info(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Custo")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Frete")
    objGridInt.colColuna.Add ("ICMS")
    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("Taxa")

   'campos de edição do grid
    objGridInt.colCampo.Add (CustoInfo.Name)
    objGridInt.colCampo.Add (DataInfo.Name)
    objGridInt.colCampo.Add (FreteInfo.Name)
    objGridInt.colCampo.Add (ICMSInfo.Name)
    objGridInt.colCampo.Add (MoedaInfo.Name)
    objGridInt.colCampo.Add (TaxaInfo.Name)

    iGrid_CustoInfo = 1
    iGrid_DataInfo = 2
    iGrid_FreteInfo = 3
    iGrid_ICMSInfo = 4
    iGrid_MoedaInfo = 5
    iGrid_TaxaInfo = 6
    
    objGridInt.objGrid = GridInfo

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 4

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 3

    GridInfo.ColWidth(0) = 900

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoIncluirNoMeioGrid = 0

    Call Grid_Inicializa(objGridInt)

    GridInfo.TextMatrix(0, 0) = "Última"
    GridInfo.TextMatrix(1, 0) = "Cotação"
    GridInfo.TextMatrix(2, 0) = "Compra"
    GridInfo.TextMatrix(3, 0) = "Entrada"

    Inicializa_Grid_Info = SUCESSO

End Function

Private Sub Opcao_Click()
    
    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(Opcao.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
            
        Select Case iFrameAtual
        
'            Case TAB_Identificacao
'                Parent.HelpContextID = IDH_ALMOXARIFADO_ID
'
'            Case TAB_Endereco
'                Parent.HelpContextID = IDH_ALMOXARIFADO_ENDERECO
                        
        End Select
    
    End If

End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Function Inicializa_Grid_ParamImp(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Valor")

   'campos de edição do grid
    objGridInt.colCampo.Add (ParamImpItem.Name)
    objGridInt.colCampo.Add (ParamImpValor.Name)

    objGridInt.objGrid = GridParamImp

    iGrid_ParamImpItem = 1
    iGrid_ParamImpValor = 2
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 50

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    GridParamImp.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ParamImp = SUCESSO

End Function

Private Function Inicializa_Grid_PlanImp(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Valor")

   'campos de edição do grid
    objGridInt.colCampo.Add (PlanImpItem.Name)
    objGridInt.colCampo.Add (PlanImpValor.Name)

    objGridInt.objGrid = GridPlanImp

    iGrid_PlanImpItem = 1
    iGrid_PlanImpValor = 2
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 50

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10

    GridPlanImp.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_PlanImp = SUCESSO

End Function

Function ParamImp_CarregaGrid() As Long

Dim lErro As Long, iLinha As Integer
Dim objMnemonicoFPTipo As ClassMnemonicoFPTipo

On Error GoTo Erro_ParamImp_CarregaGrid

    Set gcolParamPlanImp = New Collection
    
    lErro = MnemonicosPlanImpParam_Le(giFilialEmpresa, gcolParamPlanImp)
    If lErro <> SUCESSO Then gError 106696
    
    For Each objMnemonicoFPTipo In gcolParamPlanImp
    
        iLinha = iLinha + 1
        GridParamImp.TextMatrix(iLinha, iGrid_ParamImpItem) = objMnemonicoFPTipo.sMnemonico
                
    Next
    
    objGridParamImp.iLinhasExistentes = iLinha
    
    ParamImp_CarregaGrid = SUCESSO
     
    Exit Function
    
Erro_ParamImp_CarregaGrid:

    ParamImp_CarregaGrid = gErr
     
    Select Case gErr
          
        Case 106696
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158617)
     
    End Select
     
    Exit Function

End Function

Function PlanImp_CarregaGrid() As Long

Dim lErro As Long
Dim objPlanilhas As ClassPlanilhas

On Error GoTo Erro_PlanImp_CarregaGrid

    Set gcolPlanilhasImp = New Collection
    
    lErro = PlanImp_Le(giFilialEmpresa, gcolPlanilhasImp)
    If lErro <> SUCESSO Then gError 106697
    
    For Each objPlanilhas In gcolPlanilhasImp
    
        GridPlanImp.TextMatrix(objPlanilhas.iLinha, iGrid_PlanImpItem) = objPlanilhas.sTitulo
    
    Next
    
    objGridPlanImp.iLinhasExistentes = gcolPlanilhasImp.Count
    
    PlanImp_CarregaGrid = SUCESSO
     
    Exit Function
    
Erro_PlanImp_CarregaGrid:

    PlanImp_CarregaGrid = gErr
     
    Select Case gErr
          
        Case 106697
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158618)
     
    End Select
     
    Exit Function

End Function

Function PlanImp_Le(ByVal iFilialFaturamento As Integer, ByVal colPlanilhas As Collection) As Long
'carrega colPlanilhas com as linhas da planilha referente a custo de importacao
'??? falta incluir escopos mais granulares do que o geral

Dim lErro As Long, lComando As Long, iLinha As Integer, sExpressao As String, sTitulo As String
Dim objPlanilhas As ClassPlanilhas, iEscopo As Integer

On Error GoTo Erro_PlanImp_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106698
    
    sExpressao = String(STRING_FORMACAOPRECO_EXPRESSAO, 0)
    sTitulo = String(STRING_FORMACAOPRECO_TITULO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Escopo, Linha, Expressao, Titulo FROM Planilhas WHERE TipoPlanilha = ? AND FilialEmpresa = ? AND Escopo = ? ORDER BY Linha", _
        iEscopo, iLinha, sExpressao, sTitulo, PLANILHA_TIPO_IMPORTACAO, iFilialFaturamento, MNEMONICOFPRECO_ESCOPO_GERAL)
    If lErro <> AD_SQL_SUCESSO Then gError 106699
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106700
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objPlanilhas = New ClassPlanilhas
        
        With objPlanilhas
            .iTipoPlanilha = PLANILHA_TIPO_IMPORTACAO
            .iFilialEmpresa = iFilialFaturamento
            .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
            .iLinha = iLinha
            .sExpressao = sExpressao
            .sTitulo = sTitulo
        End With
        
        colPlanilhas.Add objPlanilhas
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106701
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    PlanImp_Le = SUCESSO
     
    Exit Function
    
Erro_PlanImp_Le:

    PlanImp_Le = gErr
     
    Select Case gErr
          
        Case 106699 To 106701
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_PLANILHAS", Err)
        
        Case 106698
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158619)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function MnemonicosPlanImpParam_Le(ByVal iFilialFaturamento As Integer, ByVal colMnemonicosFPTipo As Collection) As Long
'carrega colMnemonicosFPTipo com os mnemonicos que nao tenham expressao e portanto deverao ser informados pelo usuario

Dim lErro As Long, lComando As Long
Dim objMnemonicoFPTipo As ClassMnemonicoFPTipo
Dim tMnemo As typeMnemonicoFPTipo

On Error GoTo Erro_MnemonicosPlanImpParam_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106702
    
    tMnemo.sMnemonico = String(STRING_MNEMONICOFPRECO_MNEMONICO, 0)
    tMnemo.sMnemonicoDesc = String(STRING_MNEMONICOFPRECO_MNEMONICODESC, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, MnemonicoDesc FROM MnemonicoFPTIpo WHERE TipoPlanilha = ? AND FilialEmpresa = ? AND Escopo = ? AND Expressao = ? ORDER BY Mnemonico", _
        tMnemo.sMnemonico, tMnemo.sMnemonicoDesc, PLANILHA_TIPO_IMPORTACAO, iFilialFaturamento, MNEMONICOFPRECO_ESCOPO_GERAL, "")
    If lErro <> AD_SQL_SUCESSO Then gError 106703
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106704
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objMnemonicoFPTipo = New ClassMnemonicoFPTipo
    
        objMnemonicoFPTipo.sMnemonico = tMnemo.sMnemonico
        objMnemonicoFPTipo.sMnemonicoDesc = tMnemo.sMnemonicoDesc
        
        colMnemonicosFPTipo.Add objMnemonicoFPTipo
                
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106705
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    MnemonicosPlanImpParam_Le = SUCESSO
     
    Exit Function
    
Erro_MnemonicosPlanImpParam_Le:

    MnemonicosPlanImpParam_Le = gErr
     
    Select Case gErr
          
        Case 106702
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 106703 To 106705
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158620)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Sub GridParamImp_Click()
    
Dim iExecutaEntradaCelula As Integer
Dim objMnemonicoFPTipo As ClassMnemonicoFPTipo
Dim lErro As Long

    Call Grid_Click(objGridParamImp, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridParamImp, iAlterado)
        
    End If

    If GridParamImp.Row <> 0 Then
    
        'Preenche a label DescParamImp
        Set objMnemonicoFPTipo = gcolParamPlanImp.Item(GridParamImp.Row)
        
        LabelDescrParam.Caption = objMnemonicoFPTipo.sMnemonicoDesc
    
    End If
    
End Sub

Sub GridParamImp_GotFocus()

    Call Grid_Recebe_Foco(objGridParamImp)

End Sub

Sub GridParamImp_EnterCell()

    Call Grid_Entrada_Celula(objGridParamImp, iAlterado)

End Sub

Sub GridParamImp_LeaveCell()

    Call Saida_Celula(objGridParamImp)

End Sub

Sub GridParamImp_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridParamImp_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridParamImp)

    Exit Sub

Erro_GridParamImp_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158621)

    End Select

    Exit Sub

End Sub

Sub GridParamImp_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParamImp, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParamImp, iAlterado)
    End If

End Sub

Sub GridParamImp_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParamImp)
    
End Sub

Sub GridParamImp_RowColChange()

    Call Grid_RowColChange(objGridParamImp)

End Sub

Sub GridParamImp_Scroll()
    Call Grid_Scroll(objGridParamImp)
End Sub

Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz o tratamento de saida de célula

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    'Sucesso => ...
    If lErro = SUCESSO Then
        
        Select Case GridParamImp.Col

            Case iGrid_ParamImpValor
                'faz a saida da celula do valor do parametro
                lErro = Saida_Celula_Param(objGridInt)
                If lErro <> SUCESSO Then gError 116465

        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 116468
    
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr

        Case 116465
        
        Case 116468
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158622)
    
    End Select
    
    Exit Function

End Function

Public Function Saida_Celula_Param(objGridInt As AdmGrid) As Long
'faz a saida da celula percentual

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Param

    Set objGridInt.objControle = ParamImpValor
    
    'Se estiver preenchida
    If Len(Trim(ParamImpValor.Text)) > 0 Then
        
        'Critica o valor
        lErro = Valor_Critica(ParamImpValor.Text)
        If lErro <> SUCESSO Then gError 116476
        
    End If
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116477

    Saida_Celula_Param = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Param:

    Saida_Celula_Param = gErr
    
    Select Case gErr
    
        Case 116476
    
        Case 116477
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158623)
    
    End Select
    
    Exit Function

End Function

Private Sub ParamImpValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParamImpValor_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParamImpValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParamImp)
End Sub

Private Sub ParamImpValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParamImp)
End Sub

Private Sub ParamImpValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParamImp.objControle = ParamImpValor
    lErro = Grid_Campo_Libera_Foco(objGridParamImp)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Traz_CustoEmb_Tela3(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long, dCusto As Double
Dim iCondPagtoUltCot As Integer, iCondPagtoUltCom As Integer
Dim dAliquotaICMSUltCot As Double, dAliquotaICMSUltCom As Double, dAliquotaICMSUltEnt As Double
Dim dtDataUltCot As Date, dtDataUltCom As Date, dtDataUltEnt As Date
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objProdutoFilial As New ClassProdutoFilial, sMoeda As String, dTaxa As Double, dFrete As Double

On Error GoTo Erro_Traz_CustoEmb_Tela3

    lErro = CF("CustoMedioAtual_Le", objProduto.sCodigo, dCusto, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 106712
        
    LabelCustoEstoque = Format(dCusto, FORMATO_CUSTO)
    
    'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Filial
    objEstoqueProduto.sProduto = objProduto.sCodigo
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, giFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 64014 Then gError 106713
    
    LabelSaldo.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel)
    
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    objProdutoFilial.sProduto = objProduto.sCodigo
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 106714

    LabelConsumoMedio.Caption = Formata_Estoque(objProdutoFilial.dConsumoMedio)

    lErro = CF("Produto_ObterDadosUltCot", objProduto, giFilialEmpresa, dCusto, dtDataUltCot, dAliquotaICMSUltCot, sMoeda, dTaxa, dFrete)
    If lErro <> SUCESSO Then gError 106715
    
    GridInfo.TextMatrix(1, iGrid_DataInfo) = IIf(dtDataUltCot = DATA_NULA, "", Format(dtDataUltCot, "dd/mm/yyyy"))
    GridInfo.TextMatrix(1, iGrid_CustoInfo) = Format(dCusto, FORMATO_CUSTO)
    GridInfo.TextMatrix(1, iGrid_FreteInfo) = Format(dFrete, FORMATO_CUSTO)
    GridInfo.TextMatrix(1, iGrid_ICMSInfo) = CStr(Round(dAliquotaICMSUltCot * 100, 0))
    GridInfo.TextMatrix(1, iGrid_MoedaInfo) = sMoeda
    GridInfo.TextMatrix(1, iGrid_TaxaInfo) = IIf(sMoeda <> "", CStr(Round(dTaxa, 4)), "")

    lErro = CF("Produto_ObterDadosUltCompra", objProduto, giFilialEmpresa, dCusto, dtDataUltCom, dAliquotaICMSUltCom, sMoeda, dTaxa, dFrete, iCondPagtoUltCom)
    If lErro <> SUCESSO Then gError 106715
    
    GridInfo.TextMatrix(2, iGrid_DataInfo) = IIf(dtDataUltCom = DATA_NULA, "", Format(dtDataUltCom, "dd/mm/yyyy"))
    GridInfo.TextMatrix(2, iGrid_CustoInfo) = Format(dCusto, FORMATO_CUSTO)
    GridInfo.TextMatrix(2, iGrid_FreteInfo) = Format(dFrete, FORMATO_CUSTO)
    GridInfo.TextMatrix(2, iGrid_ICMSInfo) = CStr(Round(dAliquotaICMSUltCom * 100, 0))
    GridInfo.TextMatrix(2, iGrid_MoedaInfo) = sMoeda
    GridInfo.TextMatrix(2, iGrid_TaxaInfo) = IIf(sMoeda <> "", CStr(Round(dTaxa, 4)), "")

    lErro = CF("Produto_ObterDadosUltEnt", objProduto, giFilialEmpresa, dCusto, dtDataUltEnt, dAliquotaICMSUltEnt, dFrete)
    If lErro <> SUCESSO Then gError 106715
    
    GridInfo.TextMatrix(3, iGrid_DataInfo) = IIf(dtDataUltEnt = DATA_NULA, "", Format(dtDataUltEnt, "dd/mm/yyyy"))
    GridInfo.TextMatrix(3, iGrid_CustoInfo) = Format(dCusto, FORMATO_CUSTO)
    GridInfo.TextMatrix(3, iGrid_ICMSInfo) = CStr(Round(dAliquotaICMSUltEnt * 100, 0))
    GridInfo.TextMatrix(3, iGrid_FreteInfo) = Format(dFrete, FORMATO_CUSTO)
    
    If AliqNaoInf = vbChecked Then
        AliquotaICMS.Text = IIf(dtDataUltCom = DATA_NULA, IIf(dtDataUltEnt = DATA_NULA, 0, dAliquotaICMSUltEnt), dAliquotaICMSUltCom) * 100
    End If
    
    If CondPagtoNaoInf = vbChecked Then
        If iCondPagtoUltCom = 0 Then
            CondicaoPagamento.Text = ""
        Else
            CondicaoPagamento.Text = iCondPagtoUltCom
            Call CondicaoPagamento_Validate(bSGECancelDummy)
        End If
    End If
    
    If FreteNaoInf = vbChecked Then
        Frete.Text = Format(dFrete, FORMATO_CUSTO)
    End If
    
    Traz_CustoEmb_Tela3 = SUCESSO
     
    Exit Function
    
Erro_Traz_CustoEmb_Tela3:

    Traz_CustoEmb_Tela3 = gErr
     
    Select Case gErr
          
        Case 106712 To 106715
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158624)
     
    End Select
     
    Exit Function

End Function

'#########################################
'Inserido por Wagner 07/12/05
Function Preenche_ParamImp(ByVal objCustoEmbMP As ClassCustoEmbMP) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objCustoEmbMPAux As ClassCustoEmbMPAux
Dim bTodosPreechidos As Boolean

On Error GoTo Erro_Preenche_ParamImp
    
    bTodosPreechidos = True
    
    For iLinha = 1 To objGridParamImp.iLinhasExistentes
    
        Set objCustoEmbMPAux = New ClassCustoEmbMPAux
    
        objCustoEmbMPAux.iFilialEmpresa = objCustoEmbMP.iFilialEmpresa
        objCustoEmbMPAux.sMnemonico = GridParamImp.TextMatrix(iLinha, iGrid_ParamImpItem)
        objCustoEmbMPAux.sProduto = objCustoEmbMP.sProduto
        
        lErro = CF("CustoEmbMPAux_Le", objCustoEmbMPAux)
        If lErro <> SUCESSO And lErro <> 141304 Then gError 141322
        
        If lErro = SUCESSO Then
            If Len(Trim(objCustoEmbMPAux.sValor)) > 0 Then
                GridParamImp.TextMatrix(iLinha, iGrid_ParamImpValor) = objCustoEmbMPAux.sValor
            Else
                GridParamImp.TextMatrix(iLinha, iGrid_ParamImpValor) = ""
                bTodosPreechidos = False
            End If
        Else
            bTodosPreechidos = False
        End If
            
    Next
    
    If bTodosPreechidos Then Call BotaoCalcular_Click
    
    Preenche_ParamImp = SUCESSO
     
    Exit Function
    
Erro_Preenche_ParamImp:

    Preenche_ParamImp = gErr
     
    Select Case gErr
          
        Case 141322
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158625)
     
    End Select
     
    Exit Function

End Function

Function Move_ParamImp_Memoria(ByVal objCustoEmbMP As ClassCustoEmbMP) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objCustoEmbMPAux As ClassCustoEmbMPAux

On Error GoTo Erro_Move_ParamImp_Memoria

    Set objCustoEmbMP.colAux = New Collection
    
    For iLinha = 1 To objGridParamImp.iLinhasExistentes
    
        Set objCustoEmbMPAux = New ClassCustoEmbMPAux
    
        objCustoEmbMPAux.iFilialEmpresa = objCustoEmbMP.iFilialEmpresa
        objCustoEmbMPAux.sMnemonico = GridParamImp.TextMatrix(iLinha, iGrid_ParamImpItem)
        objCustoEmbMPAux.sProduto = objCustoEmbMP.sProduto
        objCustoEmbMPAux.sValor = GridParamImp.TextMatrix(iLinha, iGrid_ParamImpValor)
            
        objCustoEmbMP.colAux.Add objCustoEmbMPAux
            
    Next
    
    Move_ParamImp_Memoria = SUCESSO
     
    Exit Function
    
Erro_Move_ParamImp_Memoria:

    Move_ParamImp_Memoria = gErr
     
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158626)
     
    End Select
     
    Exit Function

End Function
'##########################################################

Private Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        Moeda.ItemData(Moeda.NewIndex) = objMoeda.iCodigo
    
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164661)
    
    End Select

End Function

Public Sub Moeda_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Moeda
End Sub

Public Sub Moeda_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Moeda
End Sub

Private Sub Trata_Moeda()

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Moeda

    If iMoedaAnt <> Codigo_Extrai(Moeda.Text) Then
    
        'Retira a mascara do produto
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116287

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            lErro = ObterMoeda(sProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    End If
    iMoedaAnt = Codigo_Extrai(Moeda.Text)

    Exit Sub

Erro_Trata_Moeda:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub

End Sub

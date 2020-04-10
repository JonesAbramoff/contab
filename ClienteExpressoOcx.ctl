VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ClienteExpressoOcx 
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5745
      Index           =   1
      Left            =   135
      TabIndex        =   55
      Top             =   630
      Width           =   9165
      Begin VB.Frame Frame11 
         Caption         =   "Tipos de Clientes"
         Height          =   1905
         Left            =   4920
         TabIndex        =   113
         Top             =   3855
         Width           =   4260
         Begin VB.CommandButton BotaoMarcarTodos 
            Height          =   480
            Index           =   0
            Left            =   3420
            Picture         =   "ClienteExpressoOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   195
            Width           =   780
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Height          =   480
            Index           =   0
            Left            =   3420
            Picture         =   "ClienteExpressoOcx.ctx":101A
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   765
            Width           =   780
         End
         Begin VB.ListBox TiposCliente 
            Height          =   1635
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   21
            Top             =   210
            Width           =   3345
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Estados"
         Height          =   1830
         Left            =   45
         TabIndex        =   111
         Top             =   -15
         Width           =   4785
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Height          =   480
            Index           =   11
            Left            =   3885
            Picture         =   "ClienteExpressoOcx.ctx":21FC
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   795
            Width           =   780
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Height          =   480
            Index           =   11
            Left            =   3885
            Picture         =   "ClienteExpressoOcx.ctx":33DE
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   225
            Width           =   780
         End
         Begin VB.ListBox UF 
            Columns         =   5
            Height          =   1410
            Left            =   75
            Style           =   1  'Checkbox
            TabIndex        =   0
            Top             =   225
            Width           =   3735
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Bairros"
         Height          =   1935
         Left            =   4920
         TabIndex        =   107
         Top             =   1890
         Width           =   4260
         Begin VB.CheckBox BairrosTodos 
            Caption         =   "Todos"
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
            Left            =   120
            TabIndex        =   18
            Top             =   210
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.Frame FrameBairro 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1470
            Left            =   120
            TabIndex        =   108
            Top             =   420
            Width           =   4095
            Begin VB.CommandButton BotaoBairros 
               Caption         =   "Bairros"
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
               Left            =   0
               TabIndex        =   20
               Top             =   1155
               Width           =   1275
            End
            Begin VB.TextBox Bairro 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   480
               MaxLength       =   250
               TabIndex        =   109
               Top             =   960
               Width           =   3075
            End
            Begin MSFlexGridLib.MSFlexGrid GridBairros 
               Height          =   1125
               Left            =   -15
               TabIndex        =   19
               Top             =   -15
               Width           =   4065
               _ExtentX        =   7170
               _ExtentY        =   1984
               _Version        =   393216
               Rows            =   6
               Cols            =   3
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cidades"
         Height          =   1920
         Left            =   4920
         TabIndex        =   66
         Top             =   -15
         Width           =   4260
         Begin VB.Frame FrameCidade 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1470
            Left            =   120
            TabIndex        =   103
            Top             =   405
            Width           =   4095
            Begin VB.TextBox Cidade 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   225
               Left            =   480
               MaxLength       =   250
               TabIndex        =   106
               Top             =   960
               Width           =   3075
            End
            Begin VB.CommandButton BotaoCidades 
               Caption         =   "Cidades"
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
               Left            =   0
               TabIndex        =   17
               Top             =   1155
               Width           =   1275
            End
            Begin MSFlexGridLib.MSFlexGrid GridCidades 
               Height          =   1125
               Left            =   -15
               TabIndex        =   16
               Top             =   0
               Width           =   4065
               _ExtentX        =   7170
               _ExtentY        =   1984
               _Version        =   393216
               Rows            =   6
               Cols            =   3
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
            End
         End
         Begin VB.CheckBox CidadesTodas 
            Caption         =   "Todas"
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
            Left            =   120
            TabIndex        =   15
            Top             =   210
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Outros"
         Height          =   2040
         Left            =   45
         TabIndex        =   65
         Top             =   2730
         Width           =   4770
         Begin VB.CheckBox UsuRespCallCenterTodos 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3765
            TabIndex        =   14
            Top             =   1680
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox UsuCobradorTodos 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3765
            TabIndex        =   12
            Top             =   1335
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox TransportadoraTodas 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3765
            TabIndex        =   10
            Top             =   945
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox VendedorTodos 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3765
            TabIndex        =   8
            Top             =   570
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox RegiaoTodas 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3765
            TabIndex        =   6
            Top             =   195
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.ComboBox UsuCobrador 
            Height          =   315
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1275
            Width           =   2355
         End
         Begin VB.ComboBox UsuRespCallCenter 
            Height          =   315
            Left            =   1335
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1650
            Width           =   2355
         End
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   900
            Width           =   2355
         End
         Begin VB.ComboBox Regiao 
            Height          =   315
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   180
            Width           =   2355
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   315
            Left            =   1335
            TabIndex        =   7
            Top             =   540
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Usu.Cobrador:"
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
            TabIndex        =   71
            Top             =   1305
            Width           =   1230
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Resp.Call C.:"
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
            Left            =   150
            TabIndex        =   70
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label TransportadoraLabel 
            AutoSize        =   -1  'True
            Caption         =   "Transport.:"
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
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   69
            Top             =   960
            Width           =   945
         End
         Begin VB.Label VendedorLabel 
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
            Left            =   390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   68
            Top             =   570
            Width           =   885
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Região:"
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
            TabIndex        =   67
            Top             =   225
            Width           =   675
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente"
         Height          =   930
         Left            =   45
         TabIndex        =   62
         Top             =   1815
         Width           =   4770
         Begin MSMask.MaskEdBox ClienteDe 
            Height          =   315
            Left            =   1335
            TabIndex        =   3
            Top             =   180
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteAte 
            Height          =   315
            Left            =   1335
            TabIndex        =   4
            Top             =   540
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelClienteAte 
            Caption         =   "Até:"
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
            TabIndex        =   64
            Top             =   585
            Width           =   435
         End
         Begin VB.Label LabelClienteDe 
            Caption         =   "De:"
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
            Left            =   960
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   63
            Top             =   225
            Width           =   360
         End
      End
      Begin VB.Frame FrameCategoriaCliente 
         Caption         =   "Categoria de Cliente"
         Height          =   975
         Left            =   45
         TabIndex        =   57
         Top             =   4785
         Width           =   4770
         Begin VB.ComboBox CategoriaClienteAte 
            Height          =   315
            Left            =   2715
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   600
            Width           =   1980
         End
         Begin VB.CheckBox CategoriaClienteTodas 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3765
            TabIndex        =   24
            Top             =   270
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.ComboBox CategoriaClienteDe 
            Height          =   315
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   1980
         End
         Begin VB.ComboBox CategoriaCliente 
            Height          =   315
            Left            =   1320
            TabIndex        =   25
            Top             =   255
            Width           =   2355
         End
         Begin VB.Label LabelCategoriaCliente 
            Caption         =   "Categoria:"
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
            Left            =   405
            TabIndex        =   61
            Top             =   300
            Width           =   855
         End
         Begin VB.Label LabelCategoriaClienteDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   60
            TabIndex        =   60
            Top             =   645
            Width           =   315
         End
         Begin VB.Label LabelCategoriaClienteAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   2370
            TabIndex        =   59
            Top             =   645
            Width           =   360
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   360
            TabIndex        =   58
            Top             =   720
            Width           =   30
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5730
      Index           =   2
      Left            =   120
      TabIndex        =   56
      Top             =   660
      Visible         =   0   'False
      Width           =   9225
      Begin VB.TextBox UFGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   4410
         TabIndex        =   112
         Top             =   1440
         Width           =   750
      End
      Begin VB.TextBox BairroGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   4575
         TabIndex        =   110
         Top             =   570
         Width           =   1785
      End
      Begin VB.CommandButton BotaoVendedorGrid 
         Caption         =   "Vendedores"
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
         Left            =   4665
         TabIndex        =   35
         Top             =   3060
         Width           =   1485
      End
      Begin VB.CommandButton BotaoHistoricoCRM 
         Caption         =   "Histórico de Contatos"
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
         Left            =   3135
         TabIndex        =   34
         Top             =   3060
         Width           =   1485
      End
      Begin VB.CommandButton BotaoHistoricoCR 
         Caption         =   "Histórico de Recebimentos"
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
         Left            =   1620
         TabIndex        =   33
         Top             =   3060
         Width           =   1485
      End
      Begin VB.CommandButton BotaoCliente 
         Caption         =   "Cliente..."
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
         Left            =   105
         TabIndex        =   32
         Top             =   3060
         Width           =   1485
      End
      Begin MSMask.MaskEdBox VendedorGrid 
         Height          =   225
         Left            =   4785
         TabIndex        =   81
         Top             =   1020
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.ComboBox UsuRespCallCenterGrid 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ClienteExpressoOcx.ctx":43F8
         Left            =   4335
         List            =   "ClienteExpressoOcx.ctx":4402
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   1845
         Width           =   1620
      End
      Begin VB.ComboBox UsuCobradorGrid 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ClienteExpressoOcx.ctx":4437
         Left            =   2115
         List            =   "ClienteExpressoOcx.ctx":4441
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   1245
         Width           =   1620
      End
      Begin VB.ComboBox TransportadoraGrid 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ClienteExpressoOcx.ctx":4476
         Left            =   2865
         List            =   "ClienteExpressoOcx.ctx":4480
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   570
         Width           =   1320
      End
      Begin VB.ComboBox RegiaoGrid 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ClienteExpressoOcx.ctx":44B5
         Left            =   1230
         List            =   "ClienteExpressoOcx.ctx":44BF
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   555
         Width           =   1635
      End
      Begin VB.TextBox CidadeGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   4500
         TabIndex        =   76
         Top             =   165
         Width           =   1785
      End
      Begin VB.TextBox FilialClienteGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3510
         TabIndex        =   75
         Top             =   165
         Width           =   945
      End
      Begin VB.TextBox ClienteGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1185
         TabIndex        =   74
         Top             =   165
         Width           =   2910
      End
      Begin VB.Frame Frame4 
         Caption         =   "Troca automática"
         Height          =   2175
         Left            =   105
         TabIndex        =   72
         Top             =   3555
         Width           =   9060
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   1500
            Index           =   5
            Left            =   165
            TabIndex        =   88
            Top             =   540
            Visible         =   0   'False
            Width           =   8685
            Begin VB.Frame Frame8 
               Caption         =   "Redistribuir"
               Height          =   1425
               Left            =   120
               TabIndex        =   98
               Top             =   15
               Width           =   8505
               Begin VB.CommandButton BotaoLimparGrid 
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
                  Height          =   330
                  Index           =   5
                  Left            =   6555
                  TabIndex        =   50
                  Top             =   225
                  Width           =   1875
               End
               Begin VB.CommandButton BotaoLimparTodos 
                  Caption         =   "Limpar Todos"
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
                  Index           =   10
                  Left            =   6555
                  TabIndex        =   51
                  Top             =   600
                  Width           =   1875
               End
               Begin VB.CommandButton BotaoAplicar 
                  Caption         =   "Aplicar nas Matrizes"
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
                  Index           =   5
                  Left            =   6555
                  TabIndex        =   52
                  Top             =   990
                  Width           =   1875
               End
               Begin VB.ComboBox CallCenterNovo 
                  Appearance      =   0  'Flat
                  Height          =   315
                  ItemData        =   "ClienteExpressoOcx.ctx":44F4
                  Left            =   675
                  List            =   "ClienteExpressoOcx.ctx":44FE
                  Style           =   2  'Dropdown List
                  TabIndex        =   101
                  Top             =   375
                  Width           =   2460
               End
               Begin MSMask.MaskEdBox CallCenterNovoPerc 
                  Height          =   315
                  Left            =   3885
                  TabIndex        =   99
                  Top             =   420
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  PromptInclude   =   0   'False
                  Enabled         =   0   'False
                  MaxLength       =   20
                  Format          =   "0%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox CallCenterNovoQtd 
                  Height          =   315
                  Left            =   2835
                  TabIndex        =   100
                  Top             =   420
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  PromptInclude   =   0   'False
                  Enabled         =   0   'False
                  MaxLength       =   6
                  Mask            =   "######"
                  PromptChar      =   " "
               End
               Begin MSFlexGridLib.MSFlexGrid GridCallCenterNovo 
                  Height          =   1155
                  Left            =   90
                  TabIndex        =   49
                  Top             =   210
                  Width           =   6330
                  _ExtentX        =   11165
                  _ExtentY        =   2037
                  _Version        =   393216
                  Rows            =   15
                  Cols            =   1
                  BackColorSel    =   -2147483643
                  ForeColorSel    =   -2147483640
                  AllowBigSelection=   0   'False
                  FocusRect       =   2
                  AllowUserResizing=   1
               End
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   1500
            Index           =   4
            Left            =   165
            TabIndex        =   87
            Top             =   540
            Visible         =   0   'False
            Width           =   8685
            Begin VB.Frame Frame7 
               Caption         =   "Redistribuir"
               Height          =   1425
               Left            =   75
               TabIndex        =   95
               Top             =   15
               Width           =   8550
               Begin VB.CommandButton BotaoLimparGrid 
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
                  Height          =   330
                  Index           =   4
                  Left            =   6585
                  TabIndex        =   46
                  Top             =   195
                  Width           =   1875
               End
               Begin VB.CommandButton BotaoLimparTodos 
                  Caption         =   "Limpar Todos"
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
                  Index           =   9
                  Left            =   6585
                  TabIndex        =   47
                  Top             =   570
                  Width           =   1875
               End
               Begin VB.CommandButton BotaoAplicar 
                  Caption         =   "Aplicar nas Matrizes"
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
                  Index           =   4
                  Left            =   6585
                  TabIndex        =   48
                  Top             =   945
                  Width           =   1875
               End
               Begin VB.ComboBox CobrNovo 
                  Appearance      =   0  'Flat
                  Height          =   315
                  ItemData        =   "ClienteExpressoOcx.ctx":4533
                  Left            =   285
                  List            =   "ClienteExpressoOcx.ctx":453D
                  Style           =   2  'Dropdown List
                  TabIndex        =   102
                  Top             =   405
                  Width           =   2580
               End
               Begin MSMask.MaskEdBox CobrNovoPerc 
                  Height          =   315
                  Left            =   3885
                  TabIndex        =   96
                  Top             =   420
                  Width           =   1440
                  _ExtentX        =   2540
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  PromptInclude   =   0   'False
                  Enabled         =   0   'False
                  MaxLength       =   20
                  Format          =   "0%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox CobrNovoQtd 
                  Height          =   315
                  Left            =   2835
                  TabIndex        =   97
                  Top             =   420
                  Width           =   1440
                  _ExtentX        =   2540
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  PromptInclude   =   0   'False
                  Enabled         =   0   'False
                  MaxLength       =   6
                  Mask            =   "######"
                  PromptChar      =   " "
               End
               Begin MSFlexGridLib.MSFlexGrid GridCobrNovo 
                  Height          =   1125
                  Left            =   75
                  TabIndex        =   45
                  Top             =   210
                  Width           =   6405
                  _ExtentX        =   11298
                  _ExtentY        =   1984
                  _Version        =   393216
                  Rows            =   15
                  Cols            =   1
                  BackColorSel    =   -2147483643
                  ForeColorSel    =   -2147483640
                  AllowBigSelection=   0   'False
                  FocusRect       =   2
                  AllowUserResizing=   1
               End
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   1500
            Index           =   3
            Left            =   165
            TabIndex        =   86
            Top             =   540
            Visible         =   0   'False
            Width           =   8685
            Begin VB.CommandButton BotaoAplicar 
               Caption         =   "Aplicar a Todos"
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
               Index           =   3
               Left            =   4845
               TabIndex        =   44
               Top             =   315
               Width           =   1875
            End
            Begin VB.ComboBox TransportadoraNova 
               Height          =   315
               Left            =   1140
               TabIndex        =   43
               Top             =   435
               Width           =   3525
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Transport.:"
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
               TabIndex        =   90
               Top             =   480
               Width           =   945
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   1500
            Index           =   2
            Left            =   165
            TabIndex        =   85
            Top             =   540
            Visible         =   0   'False
            Width           =   8685
            Begin VB.Frame Frame5 
               Caption         =   "Redistribuir"
               Height          =   1425
               Left            =   105
               TabIndex        =   91
               Top             =   30
               Width           =   8520
               Begin VB.CommandButton BotaoLimparGrid 
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
                  Height          =   255
                  Index           =   2
                  Left            =   6525
                  TabIndex        =   40
                  Top             =   495
                  Width           =   1875
               End
               Begin VB.CommandButton BotaoLimparTodos 
                  Caption         =   "Limpar Todos"
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
                  Index           =   7
                  Left            =   6525
                  TabIndex        =   41
                  Top             =   765
                  Width           =   1875
               End
               Begin VB.CommandButton BotaoVendedor 
                  Caption         =   "Vendedores"
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
                  Left            =   6525
                  TabIndex        =   39
                  Top             =   210
                  Width           =   1875
               End
               Begin VB.CommandButton BotaoAplicar 
                  Caption         =   "Aplicar a Todos"
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
                  Index           =   2
                  Left            =   6525
                  TabIndex        =   42
                  Top             =   1035
                  Width           =   1875
               End
               Begin MSMask.MaskEdBox VendNovoPerc 
                  Height          =   315
                  Left            =   3885
                  TabIndex        =   94
                  Top             =   420
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  PromptInclude   =   0   'False
                  Enabled         =   0   'False
                  MaxLength       =   20
                  Format          =   "0%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox VendNovoQtd 
                  Height          =   315
                  Left            =   2835
                  TabIndex        =   93
                  Top             =   420
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  PromptInclude   =   0   'False
                  Enabled         =   0   'False
                  MaxLength       =   6
                  Mask            =   "######"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox VendNovo 
                  Height          =   315
                  Left            =   705
                  TabIndex        =   92
                  Top             =   390
                  Width           =   2460
                  _ExtentX        =   4339
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  PromptInclude   =   0   'False
                  Enabled         =   0   'False
                  MaxLength       =   20
                  PromptChar      =   " "
               End
               Begin MSFlexGridLib.MSFlexGrid GridVendNovo 
                  Height          =   1125
                  Left            =   75
                  TabIndex        =   38
                  Top             =   210
                  Width           =   6225
                  _ExtentX        =   10980
                  _ExtentY        =   1984
                  _Version        =   393216
                  Rows            =   15
                  Cols            =   1
                  BackColorSel    =   -2147483643
                  ForeColorSel    =   -2147483640
                  AllowBigSelection=   0   'False
                  FocusRect       =   2
                  AllowUserResizing=   1
               End
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   1500
            Index           =   1
            Left            =   165
            TabIndex        =   84
            Top             =   540
            Width           =   8685
            Begin VB.CommandButton BotaoAplicar 
               Caption         =   "Aplicar a Todos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   1
               Left            =   4890
               TabIndex        =   37
               Top             =   330
               Width           =   1875
            End
            Begin VB.ComboBox RegiaoNova 
               Height          =   315
               Left            =   1155
               TabIndex        =   36
               Top             =   420
               Width           =   3630
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Região:"
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
               Left            =   345
               TabIndex        =   89
               Top             =   465
               Width           =   660
            End
         End
         Begin MSComctlLib.TabStrip TabStrip2 
            Height          =   1875
            Left            =   120
            TabIndex        =   73
            Top             =   210
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3307
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   5
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Região de Venda"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Vendedor"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Transportadora"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Cobrador"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Call Center"
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
      Begin MSFlexGridLib.MSFlexGrid GridClientes 
         Height          =   2865
         Left            =   75
         TabIndex        =   31
         Top             =   60
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   5054
         _Version        =   393216
         Rows            =   15
         Cols            =   1
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matriz:"
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
         Left            =   6270
         TabIndex        =   105
         Top             =   3210
         Width           =   585
      End
      Begin VB.Label QuantidadeMatriz 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6885
         TabIndex        =   104
         Top             =   3180
         Width           =   825
      End
      Begin VB.Label QuantidadeTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   8325
         TabIndex        =   83
         Top             =   3180
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filiais:"
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
         Left            =   7755
         TabIndex        =   82
         Top             =   3210
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7635
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   0
      Width           =   1755
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   135
         Picture         =   "ClienteExpressoOcx.ctx":4572
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   645
         Picture         =   "ClienteExpressoOcx.ctx":46CC
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1155
         Picture         =   "ClienteExpressoOcx.ctx":4BFE
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6225
      Left            =   45
      TabIndex        =   53
      Top             =   255
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10980
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filiais Clientes"
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
Attribute VB_Name = "ClienteExpressoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjClienteExpressoAnt As ClassClienteExpressoSel

Dim gbTrazendoDados As Boolean
Dim gbLimpandoDados As Boolean

Dim giClienteInicial As Integer
Dim giVendedor As Integer

Const VENDEDOR_CLI = 1
Const VENDEDOR_FIL = 2
Const VENDEDOR_NOV = 3

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Private WithEvents objEventoBairro As AdmEvento
Attribute objEventoBairro.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim iFiltroAlterado As Integer
Dim iFrameAtual As Integer
Dim iFrame2Atual As Integer

Dim objGridCidades As AdmGrid
Dim iGrid_Cidade_Col As Integer

Dim objGridBairros As AdmGrid
Dim iGrid_Bairro_Col As Integer

Dim objGridCobrNovo As AdmGrid
Dim iGrid_CobrNovo_Col As Integer
Dim iGrid_CobrNovoQtd_Col As Integer
Dim iGrid_CobrNovoPerc_Col As Integer

Dim objGridCallCenterNovo As AdmGrid
Dim iGrid_CallCenterNovo_Col As Integer
Dim iGrid_CallCenterNovoQtd_Col As Integer
Dim iGrid_CallCenterNovoPerc_Col As Integer

Dim objGridVendNovo As AdmGrid
Dim iGrid_VendNovo_Col As Integer
Dim iGrid_VendNovoQtd_Col As Integer
Dim iGrid_VendNovoPerc_Col As Integer

Dim objGridClientes As AdmGrid
Dim iGrid_Cliente_Col As Integer
Dim iGrid_FilialCliente_Col As Integer
Dim iGrid_CliUF_Col As Integer
Dim iGrid_CliCidade_Col As Integer
Dim iGrid_CliBairro_Col As Integer
Dim iGrid_Regiao_Col As Integer
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_Transp_Col As Integer
Dim iGrid_Cobr_Col As Integer
Dim iGrid_CallCenter_Col As Integer

Const TAB_Selecao = 1
Const TAB_CLIENTE = 2

Const FRAME2_REGIAO = 1
Const FRAME2_VEND = 2
Const FRAME2_TRANSP = 3
Const FRAME2_COBR = 4
Const FRAME2_CALLCENTER = 5

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cliente - Expresso"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ClienteExpresso"

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

Private Sub TiposCliente_Click()
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UF_Click()
    iFiltroAlterado = REGISTRO_ALTERADO
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
    'Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    'gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload
       
    Call ComandoSeta_Liberar(Me.Name)

    Set objGridClientes = Nothing
    Set objGridCallCenterNovo = Nothing
    Set objGridCidades = Nothing
    Set objGridBairros = Nothing
    Set objGridCobrNovo = Nothing
    Set objGridVendNovo = Nothing
    
    Set objEventoCliente = Nothing
    Set objEventoCidade = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoBairro = Nothing
    
    Set gobjClienteExpressoAnt = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202031)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Form_Load

    gbTrazendoDados = True
    gbLimpandoDados = False

    Set objEventoCliente = New AdmEvento
    Set objEventoCidade = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoBairro = New AdmEvento

    Set objGridClientes = New AdmGrid
    Set objGridCallCenterNovo = New AdmGrid
    Set objGridCidades = New AdmGrid
    Set objGridBairros = New AdmGrid
    Set objGridCobrNovo = New AdmGrid
    Set objGridVendNovo = New AdmGrid
       
    Set gobjClienteExpressoAnt = New ClassClienteExpressoSel
    
    lErro = Inicializa_GridClientes(objGridClientes)
    If lErro <> SUCESSO Then gError 202032

    lErro = Inicializa_GridCallCenterNovo(objGridCallCenterNovo)
    If lErro <> SUCESSO Then gError 202033

    lErro = Inicializa_GridCidades(objGridCidades)
    If lErro <> SUCESSO Then gError 202034

    lErro = Inicializa_GridBairros(objGridBairros)
    If lErro <> SUCESSO Then gError 202034

    lErro = Inicializa_GridCobrNovo(objGridCobrNovo)
    If lErro <> SUCESSO Then gError 202035

    lErro = Inicializa_GridVendNovo(objGridVendNovo)
    If lErro <> SUCESSO Then gError 202036

    lErro = Carrega_ComboCategoriaCliente(CategoriaCliente)
    If lErro <> SUCESSO Then gError 202037

    lErro = Carrega_Usuarios()
    If lErro <> SUCESSO Then gError 202038
    
    lErro = Carrega_Regiao()
    If lErro <> SUCESSO Then gError 202039
    
    lErro = Carrega_Transportadora()
    If lErro <> SUCESSO Then gError 202040
    
    lErro = Carrega_Estados()
    If lErro <> SUCESSO Then gError 202040
    
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteDe.Enabled = False
    CategoriaClienteAte.Enabled = False
    CategoriaClienteDe.ListIndex = -1
    CategoriaClienteAte.ListIndex = -1
    
    TransportadoraTodas.Value = vbChecked
    Call TransportadoraTodas_Click
    RegiaoTodas.Value = vbChecked
    Call RegiaoTodas_Click
    VendedorTodos.Value = vbChecked
    Call VendedorTodos_Click
    UsuCobradorTodos.Value = vbChecked
    Call UsuCobradorTodos_Click
    UsuRespCallCenterTodos.Value = vbChecked
    Call UsuRespCallCenterTodos_Click
    CidadesTodas.Value = vbChecked
    Call CidadesTodas_Click
    BairrosTodos.Value = vbChecked
    Call BairrosTodos_Click
    
    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", "Codigo", "Descricao", 255, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 200224
    
    For Each objCodigoNome In colCodigoDescricao
        TiposCliente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        TiposCliente.ItemData(TiposCliente.NewIndex) = objCodigoNome.iCodigo
        TiposCliente.Selected(TiposCliente.NewIndex) = True
    Next

    iFrameAtual = TAB_Selecao
    iFrame2Atual = FRAME2_REGIAO
    iAlterado = 0
    iFiltroAlterado = REGISTRO_ALTERADO

    lErro_Chama_Tela = SUCESSO
   
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 202032 To 202040
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202041)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202042)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Selecao_Memoria(ByVal objClienteExpresso As ClassClienteExpressoSel) As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Move_Selecao_Memoria
   
    objClienteExpresso.sUsuCobrador = UsuCobrador.Text
    objClienteExpresso.sUsuRespCallCenter = UsuRespCallCenter.Text
    
    If CategoriaClienteTodas.Value = vbChecked Then
        objClienteExpresso.sCategoria = ""
        objClienteExpresso.sCategoriaDe = ""
        objClienteExpresso.sCategoriaAte = ""
    Else
        If CategoriaCliente.Text = "" Then gError 202042
        
        objClienteExpresso.sCategoria = CategoriaCliente.Text
        objClienteExpresso.sCategoriaDe = CategoriaClienteDe.Text
        objClienteExpresso.sCategoriaAte = CategoriaClienteAte.Text
    End If
    
    objClienteExpresso.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    objClienteExpresso.iRegiao = Codigo_Extrai(Regiao.Text)
    objClienteExpresso.iVendedor = Codigo_Extrai(Vendedor.Text)
    objClienteExpresso.lClienteDe = LCodigo_Extrai(ClienteDe.Text)
    objClienteExpresso.lClienteAte = LCodigo_Extrai(ClienteAte.Text)
    
    'Muda o vazio de acordo com o campos todos,
    'se todos estiver desmarcado o campo vazio assume o significado de trazer os registros com o campo não preenchido
    If objClienteExpresso.iCodTransportadora = 0 And TransportadoraTodas.Value = vbChecked Then objClienteExpresso.iCodTransportadora = -1
    If objClienteExpresso.iRegiao = 0 And RegiaoTodas.Value = vbChecked Then objClienteExpresso.iRegiao = -1
    If objClienteExpresso.iVendedor = 0 And VendedorTodos.Value = vbChecked Then objClienteExpresso.iVendedor = -1
    If objClienteExpresso.sUsuCobrador = "" And UsuCobradorTodos.Value = vbChecked Then objClienteExpresso.sUsuCobrador = "-1"
    If objClienteExpresso.sUsuRespCallCenter = "" And UsuRespCallCenterTodos.Value = vbChecked Then objClienteExpresso.sUsuRespCallCenter = "-1"
    
    For iLinha = 1 To objGridCidades.iLinhasExistentes
        If Len(Trim(GridCidades.TextMatrix(iLinha, iGrid_Cidade_Col))) > 0 Then
            objClienteExpresso.colCidades.Add GridCidades.TextMatrix(iLinha, iGrid_Cidade_Col)
        End If
    Next
    
    For iLinha = 1 To objGridBairros.iLinhasExistentes
        If Len(Trim(GridBairros.TextMatrix(iLinha, iGrid_Bairro_Col))) > 0 Then
            objClienteExpresso.colBairros.Add GridBairros.TextMatrix(iLinha, iGrid_Bairro_Col)
        End If
    Next
    
    For iLinha = 0 To UF.ListCount - 1
        If UF.Selected(iLinha) Then
            objClienteExpresso.colUFs.Add UF.List(iLinha)
        End If
    Next
    
    For iLinha = 0 To TiposCliente.ListCount - 1
        If TiposCliente.Selected(iLinha) Then
            objClienteExpresso.colTipoCli.Add Codigo_Extrai(TiposCliente.List(iLinha))
        End If
    Next
    
    If objClienteExpresso.colTipoCli.Count = TiposCliente.ListCount Then objClienteExpresso.iTodosTipoCli = MARCADO
    
    If objClienteExpresso.colUFs.Count = UF.ListCount Then objClienteExpresso.iTodasUFs = MARCADO
    
    If CidadesTodas.Value = vbUnchecked And objClienteExpresso.colCidades.Count = 0 Then gError 202043
    
    If BairrosTodos.Value = vbUnchecked And objClienteExpresso.colBairros.Count = 0 Then gError 202144
    
    If objClienteExpresso.colUFs.Count = 0 Then gError 202146
    If objClienteExpresso.colTipoCli.Count = 0 Then gError 202147
    
    If objClienteExpresso.sCategoriaDe > objClienteExpresso.sCategoriaAte Then gError 202044
    
    If objClienteExpresso.lClienteDe > objClienteExpresso.lClienteAte Then gError 202045
   
    Move_Selecao_Memoria = SUCESSO

    Exit Function

Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr

    Select Case gErr
           
        Case 202042
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_INFORMADA", gErr)
            
        Case 202043
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_CIDADE_PREENCHIDA", gErr)
            
        Case 202044
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_ITEM_INICIAL_MAIOR", gErr)
            
        Case 202045
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            
        Case 202144
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_BAIRRO_PREENCHIDO", gErr)
            
        Case 202146
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ESTADO_PREENCHIDO", gErr)
        
        Case 202147
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_TIPO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202046)

    End Select

    Exit Function

End Function

Function Trata_Selecao(ByVal objClienteExpresso As ClassClienteExpressoSel) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Selecao

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("ClienteExpresso_Le", objClienteExpresso)
    If lErro <> SUCESSO Then gError 202047
    
    If objClienteExpresso.colClientes.Count = 0 Then gError 202048
    
    lErro = Preenche_GridCliente(objClienteExpresso)
    If lErro <> SUCESSO Then gError 202049
    
    GL_objMDIForm.MousePointer = vbDefault
   
    Trata_Selecao = SUCESSO

    Exit Function

Erro_Trata_Selecao:

    GL_objMDIForm.MousePointer = vbDefault

    Trata_Selecao = gErr

    Select Case gErr
    
        Case 202047, 202049
        
        Case 202048
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_COBRANCA_SEM_CLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202050)

    End Select

    Exit Function

End Function

Function Preenche_GridCliente(ByVal objClienteExpresso As ClassClienteExpressoSel) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objcliente As ClassCliente
Dim objFilial As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim objVendedor As ClassVendedor
Dim iMatriz As Integer

On Error GoTo Erro_Preenche_GridCliente

    Call Grid_Limpa(objGridClientes)
    Call Grid_Limpa(objGridCobrNovo)
    Call Grid_Limpa(objGridVendNovo)
    Call Grid_Limpa(objGridCallCenterNovo)
    
    iAlterado = 0
    
    'Aumenta o número de linhas do grid se necessário
    If objClienteExpresso.colClientes.Count >= objGridClientes.objGrid.Rows Then
        Call Refaz_Grid(objGridClientes, objClienteExpresso.colClientes.Count)
    End If

    iIndice = 0
    For Each objcliente In objClienteExpresso.colClientes
    
        iIndice = iIndice + 1
        Set objFilial = objClienteExpresso.colFiliais.Item(iIndice)
        Set objEndereco = New ClassEndereco
        Set objVendedor = New ClassVendedor
        
        objEndereco.lCodigo = objFilial.lEndereco
        objVendedor.iCodigo = objFilial.iVendedor
   
        GridClientes.TextMatrix(iIndice, iGrid_Cliente_Col) = objcliente.lCodigo & SEPARADOR & objcliente.sNomeReduzido
        GridClientes.TextMatrix(iIndice, iGrid_FilialCliente_Col) = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
        
        If objFilial.iCodFilial = FILIAL_MATRIZ Then
            iMatriz = iMatriz + 1
        End If
        
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 202051
        GridClientes.TextMatrix(iIndice, iGrid_CliCidade_Col) = objEndereco.sCidade
        GridClientes.TextMatrix(iIndice, iGrid_CliBairro_Col) = objEndereco.sBairro
        GridClientes.TextMatrix(iIndice, iGrid_CliUF_Col) = objEndereco.sSiglaEstado
        
        If objFilial.iRegiao <> 0 Then
            Call Combo_Seleciona_ItemData(RegiaoGrid, objFilial.iRegiao)
            GridClientes.TextMatrix(iIndice, iGrid_Regiao_Col) = RegiaoGrid.Text
        End If

        If objVendedor.iCodigo <> 0 Then
            VendedorGrid.Text = CStr(objVendedor.iCodigo)
            lErro = TP_Vendedor_Le2(VendedorGrid, objVendedor, DESMARCADO)
            If lErro <> SUCESSO Then gError 202052
            GridClientes.TextMatrix(iIndice, iGrid_Vendedor_Col) = VendedorGrid.Text
        End If
        
        If objFilial.iCodTransportadora <> 0 Then
            Call Combo_Seleciona_ItemData(TransportadoraGrid, objFilial.iCodTransportadora)
            GridClientes.TextMatrix(iIndice, iGrid_Transp_Col) = TransportadoraGrid.Text
        End If
        
        GridClientes.TextMatrix(iIndice, iGrid_Cobr_Col) = objcliente.sUsuarioCobrador
        GridClientes.TextMatrix(iIndice, iGrid_CallCenter_Col) = objcliente.sUsuRespCallCenter
    
    Next
           
    objGridClientes.iLinhasExistentes = iIndice
    
    QuantidadeTotal.Caption = CStr(iIndice)
    QuantidadeMatriz.Caption = CStr(iMatriz)
        
    Call Ordenacao_Limpa(objGridClientes)
    
    Preenche_GridCliente = SUCESSO

    Exit Function

Erro_Preenche_GridCliente:

    Preenche_GridCliente = gErr

    Select Case gErr
    
        Case 202051, 202052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202053)

    End Select

    Exit Function

End Function

Private Sub TabStrip2_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip2)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
Dim objClienteExpresso As New ClassClienteExpressoSel

On Error GoTo Erro_TabStrip1_BeforeClick

    gbTrazendoDados = True

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
    
    'Se estava no tab de seleção e está passando para outro tab
    If iFrameAtual = TAB_Selecao And Not gbLimpandoDados Then
    
        'Valida a seleção
        lErro = Move_Selecao_Memoria(objClienteExpresso)
        If lErro <> SUCESSO Then gError 202054
        
        If objClienteExpresso.iCodTransportadora <> gobjClienteExpressoAnt.iCodTransportadora Or _
            objClienteExpresso.iRegiao <> gobjClienteExpressoAnt.iRegiao Or _
            objClienteExpresso.iVendedor <> gobjClienteExpressoAnt.iVendedor Or _
            objClienteExpresso.lClienteAte <> gobjClienteExpressoAnt.lClienteAte Or _
            objClienteExpresso.lClienteDe <> gobjClienteExpressoAnt.lClienteDe Or _
            objClienteExpresso.sCategoria <> gobjClienteExpressoAnt.sCategoria Or _
            objClienteExpresso.sCategoriaAte <> gobjClienteExpressoAnt.sCategoriaAte Or _
            objClienteExpresso.sCategoriaDe <> gobjClienteExpressoAnt.sCategoriaDe Or _
            objClienteExpresso.sUsuCobrador <> gobjClienteExpressoAnt.sUsuCobrador Or _
            objClienteExpresso.sUsuRespCallCenter <> gobjClienteExpressoAnt.sUsuRespCallCenter Or _
            iFiltroAlterado = REGISTRO_ALTERADO Then
        
            lErro = Trata_Selecao(objClienteExpresso)
            If lErro <> SUCESSO Then gError 202055
            
            Set gobjClienteExpressoAnt = objClienteExpresso
            iFiltroAlterado = 0
            
        End If
    
    End If
    
    gbTrazendoDados = False

    Exit Sub

Erro_TabStrip1_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 202054, 202055

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202056)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip2_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index <> iFrame2Atual Then

        If TabStrip_PodeTrocarTab(iFrame2Atual, TabStrip2, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame2(TabStrip2.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame2(iFrame2Atual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrame2Atual = TabStrip2.SelectedItem.Index
        
    End If

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

Private Function Inicializa_GridClientes(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    objGrid.colColuna.Add ("UF")
    objGrid.colColuna.Add ("Cidade")
    objGrid.colColuna.Add ("Bairro")
    objGrid.colColuna.Add ("Região")
    objGrid.colColuna.Add ("Vendedor")
    objGrid.colColuna.Add ("Transportadora")
    objGrid.colColuna.Add ("Usu. Cobrador")
    objGrid.colColuna.Add ("Usu. Call Center")
       
    'Controles que participam do Grid
    objGrid.colCampo.Add (ClienteGrid.Name)
    objGrid.colCampo.Add (FilialClienteGrid.Name)
    objGrid.colCampo.Add (UFGrid.Name)
    objGrid.colCampo.Add (CidadeGrid.Name)
    objGrid.colCampo.Add (BairroGrid.Name)
    objGrid.colCampo.Add (RegiaoGrid.Name)
    objGrid.colCampo.Add (VendedorGrid.Name)
    objGrid.colCampo.Add (TransportadoraGrid.Name)
    objGrid.colCampo.Add (UsuCobradorGrid.Name)
    objGrid.colCampo.Add (UsuRespCallCenterGrid.Name)

    'Colunas do Grid
    iGrid_Cliente_Col = 1
    iGrid_FilialCliente_Col = 2
    iGrid_CliUF_Col = 3
    iGrid_CliCidade_Col = 4
    iGrid_CliBairro_Col = 5
    iGrid_Regiao_Col = 6
    iGrid_Vendedor_Col = 7
    iGrid_Transp_Col = 8
    iGrid_Cobr_Col = 9
    iGrid_CallCenter_Col = 10

    objGrid.objGrid = GridClientes
    
    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridClientes.ColWidth(0) = 600

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridClientes = SUCESSO

End Function

Private Sub GridClientes_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridClientes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridClientes, iAlterado)
    End If
    
    colcolColecoes.Add gobjClienteExpressoAnt.colClientes
    colcolColecoes.Add gobjClienteExpressoAnt.colFiliais
    
    If Not gbTrazendoDados Then Call Ordenacao_ClickGrid(objGridClientes, , colcolColecoes)

End Sub

Private Sub GridClientes_GotFocus()
    Call Grid_Recebe_Foco(objGridClientes)
End Sub

Private Sub GridClientes_EnterCell()
    Call Grid_Entrada_Celula(objGridClientes, iAlterado)
End Sub

Private Sub GridClientes_LeaveCell()
    Call Saida_Celula(objGridClientes)
End Sub

Private Sub GridClientes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridClientes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridClientes, iAlterado)
    End If

End Sub

Private Sub GridClientes_RowColChange()

Dim iLinhaAnt As Integer

    Call Grid_RowColChange(objGridClientes)
    
    If Not gbTrazendoDados Then
        If objGridClientes.iLinhaAntiga <> objGridClientes.objGrid.Row Then
            objGridClientes.iLinhaAntiga = objGridClientes.objGrid.Row
        End If
    End If
    
End Sub

Private Sub GridClientes_Scroll()
    Call Grid_Scroll(objGridClientes)
End Sub

Private Sub GridClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridClientes)
End Sub

Private Sub GridClientes_LostFocus()
    Call Grid_Libera_Foco(objGridClientes)
End Sub

Private Function Inicializa_GridCidades(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Cidade")
       
    'Controles que participam do Grid
    objGrid.colCampo.Add (Cidade.Name)

    'Colunas do Grid
    iGrid_Cidade_Col = 1

    objGrid.objGrid = GridCidades
    
    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridCidades.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridCidades = SUCESSO

End Function

Private Sub GridCidades_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCidades, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCidades, iAlterado)
    End If

End Sub

Private Sub GridCidades_GotFocus()
    Call Grid_Recebe_Foco(objGridCidades)
End Sub

Private Sub GridCidades_EnterCell()
    Call Grid_Entrada_Celula(objGridCidades, iAlterado)
End Sub

Private Sub GridCidades_LeaveCell()
    Call Saida_Celula(objGridCidades)
End Sub

Private Sub GridCidades_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCidades, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCidades, iAlterado)
    End If

End Sub

Private Sub GridCidades_RowColChange()

Dim iLinhaAnt As Integer

    Call Grid_RowColChange(objGridCidades)
    
    If Not gbTrazendoDados Then
        If objGridCidades.iLinhaAntiga <> objGridCidades.objGrid.Row Then
            objGridCidades.iLinhaAntiga = objGridCidades.objGrid.Row
        End If
    End If
    
End Sub

Private Sub GridCidades_Scroll()
    Call Grid_Scroll(objGridCidades)
End Sub

Private Sub GridCidades_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCidades)
End Sub

Private Sub GridCidades_LostFocus()
    Call Grid_Libera_Foco(objGridCidades)
End Sub

Private Function Inicializa_GridBairros(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Bairro")
       
    'Controles que participam do Grid
    objGrid.colCampo.Add (Bairro.Name)

    'Colunas do Grid
    iGrid_Bairro_Col = 1

    objGrid.objGrid = GridBairros
    
    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridBairros.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridBairros = SUCESSO

End Function

Private Sub GridBairros_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBairros, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBairros, iAlterado)
    End If

End Sub

Private Sub GridBairros_GotFocus()
    Call Grid_Recebe_Foco(objGridBairros)
End Sub

Private Sub GridBairros_EnterCell()
    Call Grid_Entrada_Celula(objGridBairros, iAlterado)
End Sub

Private Sub GridBairros_LeaveCell()
    Call Saida_Celula(objGridBairros)
End Sub

Private Sub GridBairros_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBairros, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBairros, iAlterado)
    End If

End Sub

Private Sub GridBairros_RowColChange()

Dim iLinhaAnt As Integer

    Call Grid_RowColChange(objGridBairros)
    
    If Not gbTrazendoDados Then
        If objGridBairros.iLinhaAntiga <> objGridBairros.objGrid.Row Then
            objGridBairros.iLinhaAntiga = objGridBairros.objGrid.Row
        End If
    End If
    
End Sub

Private Sub GridBairros_Scroll()
    Call Grid_Scroll(objGridBairros)
End Sub

Private Sub GridBairros_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridBairros)
End Sub

Private Sub GridBairros_LostFocus()
    Call Grid_Libera_Foco(objGridBairros)
End Sub

Private Function Inicializa_GridCobrNovo(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Cobrador")
    objGrid.colColuna.Add ("Qtde")
    objGrid.colColuna.Add ("%")
       
    'Controles que participam do Grid
    objGrid.colCampo.Add (CobrNovo.Name)
    objGrid.colCampo.Add (CobrNovoQtd.Name)
    objGrid.colCampo.Add (CobrNovoPerc.Name)

    'Colunas do Grid
    iGrid_CobrNovo_Col = 1
    iGrid_CobrNovoQtd_Col = 2
    iGrid_CobrNovoPerc_Col = 3

    objGrid.objGrid = GridCobrNovo

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridCobrNovo.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridCobrNovo = SUCESSO

End Function

Private Sub GridCobrNovo_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCobrNovo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCobrNovo, iAlterado)
    End If

End Sub

Private Sub GridCobrNovo_GotFocus()
    Call Grid_Recebe_Foco(objGridCobrNovo)
End Sub

Private Sub GridCobrNovo_EnterCell()
    Call Grid_Entrada_Celula(objGridCobrNovo, iAlterado)
End Sub

Private Sub GridCobrNovo_LeaveCell()
    Call Saida_Celula(objGridCobrNovo)
End Sub

Private Sub GridCobrNovo_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCobrNovo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCobrNovo, iAlterado)
    End If

End Sub

Private Sub GridCobrNovo_RowColChange()

Dim iLinhaAnt As Integer

    Call Grid_RowColChange(objGridCobrNovo)
    
    If Not gbTrazendoDados Then
        If objGridCobrNovo.iLinhaAntiga <> objGridCobrNovo.objGrid.Row Then
            objGridCobrNovo.iLinhaAntiga = objGridCobrNovo.objGrid.Row
        End If
    End If
    
End Sub

Private Sub GridCobrNovo_Scroll()
    Call Grid_Scroll(objGridCobrNovo)
End Sub

Private Sub GridCobrNovo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCobrNovo)
End Sub

Private Sub GridCobrNovo_LostFocus()
    Call Grid_Libera_Foco(objGridCobrNovo)
End Sub

Private Function Inicializa_GridVendNovo(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Vendedor")
    objGrid.colColuna.Add ("Qtde")
    objGrid.colColuna.Add ("%")
       
    'Controles que participam do Grid
    objGrid.colCampo.Add (VendNovo.Name)
    objGrid.colCampo.Add (VendNovoQtd.Name)
    objGrid.colCampo.Add (VendNovoPerc.Name)

    'Colunas do Grid
    iGrid_VendNovo_Col = 1
    iGrid_VendNovoQtd_Col = 2
    iGrid_VendNovoPerc_Col = 3

    objGrid.objGrid = GridVendNovo

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridVendNovo.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridVendNovo = SUCESSO

End Function

Private Sub GridVendNovo_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridVendNovo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVendNovo, iAlterado)
    End If

End Sub

Private Sub GridVendNovo_GotFocus()
    Call Grid_Recebe_Foco(objGridVendNovo)
End Sub

Private Sub GridVendNovo_EnterCell()
    Call Grid_Entrada_Celula(objGridVendNovo, iAlterado)
End Sub

Private Sub GridVendNovo_LeaveCell()
    Call Saida_Celula(objGridVendNovo)
End Sub

Private Sub GridVendNovo_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridVendNovo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVendNovo, iAlterado)
    End If

End Sub

Private Sub GridVendNovo_RowColChange()

Dim iLinhaAnt As Integer

    Call Grid_RowColChange(objGridVendNovo)
    
    If Not gbTrazendoDados Then
        If objGridVendNovo.iLinhaAntiga <> objGridVendNovo.objGrid.Row Then
            objGridVendNovo.iLinhaAntiga = objGridVendNovo.objGrid.Row
        End If
    End If
    
End Sub

Private Sub GridVendNovo_Scroll()
    Call Grid_Scroll(objGridVendNovo)
End Sub

Private Sub GridVendNovo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridVendNovo)
End Sub

Private Sub GridVendNovo_LostFocus()
    Call Grid_Libera_Foco(objGridVendNovo)
End Sub

Private Function Inicializa_GridCallCenterNovo(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Responsável")
    objGrid.colColuna.Add ("Qtde")
    objGrid.colColuna.Add ("%")
       
    'Controles que participam do Grid
    objGrid.colCampo.Add (CallCenterNovo.Name)
    objGrid.colCampo.Add (CallCenterNovoQtd.Name)
    objGrid.colCampo.Add (CallCenterNovoPerc.Name)

    'Colunas do Grid
    iGrid_CallCenterNovo_Col = 1
    iGrid_CallCenterNovoQtd_Col = 2
    iGrid_CallCenterNovoPerc_Col = 3

    objGrid.objGrid = GridCallCenterNovo
    
    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridCallCenterNovo.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridCallCenterNovo = SUCESSO

End Function

Private Sub GridCallCenterNovo_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCallCenterNovo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCallCenterNovo, iAlterado)
    End If

End Sub

Private Sub GridCallCenterNovo_GotFocus()
    Call Grid_Recebe_Foco(objGridCallCenterNovo)
End Sub

Private Sub GridCallCenterNovo_EnterCell()
    Call Grid_Entrada_Celula(objGridCallCenterNovo, iAlterado)
End Sub

Private Sub GridCallCenterNovo_LeaveCell()
    Call Saida_Celula(objGridCallCenterNovo)
End Sub

Private Sub GridCallCenterNovo_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCallCenterNovo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCallCenterNovo, iAlterado)
    End If

End Sub

Private Sub GridCallCenterNovo_RowColChange()

Dim iLinhaAnt As Integer

    Call Grid_RowColChange(objGridCallCenterNovo)
    
    If Not gbTrazendoDados Then
        If objGridCallCenterNovo.iLinhaAntiga <> objGridCallCenterNovo.objGrid.Row Then
            objGridCallCenterNovo.iLinhaAntiga = objGridCallCenterNovo.objGrid.Row
        End If
    End If
    
End Sub

Private Sub GridCallCenterNovo_Scroll()
    Call Grid_Scroll(objGridCallCenterNovo)
End Sub

Private Sub GridCallCenterNovo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCallCenterNovo)
End Sub

Private Sub GridCallCenterNovo_LostFocus()
    Call Grid_Libera_Foco(objGridCallCenterNovo)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        'Clientes
        If objGridInt.objGrid.Name = GridClientes.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Regiao_Col
                    lErro = Saida_Celula_Padrao(objGridInt, RegiaoGrid)
                    If lErro <> SUCESSO Then gError 202057

                Case iGrid_Vendedor_Col
                    lErro = Saida_Celula_Vendedor(objGridInt, VendedorGrid)
                    If lErro <> SUCESSO Then gError 202058
                    
                Case iGrid_Transp_Col
                    lErro = Saida_Celula_Padrao(objGridInt, TransportadoraGrid)
                    If lErro <> SUCESSO Then gError 202059

                Case iGrid_Cobr_Col
                    lErro = Saida_Celula_Padrao(objGridInt, UsuCobradorGrid)
                    If lErro <> SUCESSO Then gError 202060

                Case iGrid_CallCenter_Col
                    lErro = Saida_Celula_Padrao(objGridInt, UsuRespCallCenterGrid)
                    If lErro <> SUCESSO Then gError 202061

            End Select
            
        ElseIf objGridInt.objGrid.Name = GridCidades.Name Then
                    
            Select Case objGridInt.objGrid.Col

                Case iGrid_Cidade_Col
                    lErro = Saida_Celula_Cidade(objGridInt, True)
                    If lErro <> SUCESSO Then gError 202062
                    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridBairros.Name Then
                    
            Select Case objGridInt.objGrid.Col

                Case iGrid_Bairro_Col
                    lErro = Saida_Celula_Padrao(objGridInt, Bairro, True, True)
                    If lErro <> SUCESSO Then gError 202062
                    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridCobrNovo.Name Then
                    
            Select Case objGridInt.objGrid.Col

                Case iGrid_CobrNovo_Col
                    lErro = Saida_Celula_Padrao(objGridInt, CobrNovo, True, True)
                    If lErro <> SUCESSO Then gError 202063
                    
                Case iGrid_CobrNovoQtd_Col
                    lErro = Saida_Celula_Quantidade(objGridInt, CobrNovoQtd)
                    If lErro <> SUCESSO Then gError 202064
                    
                Case iGrid_CobrNovoPerc_Col
                    lErro = Saida_Celula_Percentual(objGridInt, CobrNovoPerc)
                    If lErro <> SUCESSO Then gError 202065
                    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridVendNovo.Name Then
                    
            Select Case objGridInt.objGrid.Col

                Case iGrid_VendNovo_Col
                    lErro = Saida_Celula_Vendedor(objGridInt, VendNovo, True, True)
                    If lErro <> SUCESSO Then gError 202066
                    
                Case iGrid_VendNovoQtd_Col
                    lErro = Saida_Celula_Quantidade(objGridInt, VendNovoQtd)
                    If lErro <> SUCESSO Then gError 202067
                    
                Case iGrid_VendNovoPerc_Col
                    lErro = Saida_Celula_Percentual(objGridInt, VendNovoPerc)
                    If lErro <> SUCESSO Then gError 202068
                    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridCallCenterNovo.Name Then
                     
            Select Case objGridInt.objGrid.Col

                Case iGrid_CallCenterNovo_Col
                    lErro = Saida_Celula_Padrao(objGridInt, CallCenterNovo, True, True)
                    If lErro <> SUCESSO Then gError 202069
                    
                Case iGrid_CallCenterNovoQtd_Col
                    lErro = Saida_Celula_Quantidade(objGridInt, CallCenterNovoQtd)
                    If lErro <> SUCESSO Then gError 202070
                    
                Case iGrid_CallCenterNovoPerc_Col
                    lErro = Saida_Celula_Percentual(objGridInt, CallCenterNovoPerc)
                    If lErro <> SUCESSO Then gError 202071
                    
            End Select
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 202072

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 202057 To 202071

        Case 202072
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202073)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bAdicionaLinha As Boolean = False, Optional ByVal bTestaRepeticao As Boolean = False) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinhasAnt As Integer
Dim iIndice As Integer
Dim iQtd As Integer

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        iLinhasAnt = objGridInt.iLinhasExistentes
        
        If bTestaRepeticao Then
            For iIndice = 1 To objGridInt.iLinhasExistentes
                If iIndice <> objGridInt.objGrid.Row Then
                    If UCase(objControle.Text) = UCase(objGridInt.objGrid.TextMatrix(iIndice, objGridInt.objGrid.Col)) Then gError 202145
                End If
            Next
        End If
       
        If bAdicionaLinha Then
            Call Adiciona_Linha(objGridInt)
        End If
        
        If (Me.ActiveControl Is CobrNovo Or Me.ActiveControl Is CallCenterNovo) And iLinhasAnt <> objGridInt.iLinhasExistentes Then
            
            iQtd = 0
            For iIndice = 1 To objGridInt.objGrid.Row - 1
                iQtd = iQtd + StrParaInt(objGridInt.objGrid.TextMatrix(iIndice, 2))
            Next
            
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, 2) = CStr(StrParaInt(QuantidadeMatriz.Caption) - iQtd)
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, 3) = Format((StrParaInt(QuantidadeMatriz.Caption) - iQtd) / StrParaInt(QuantidadeMatriz.Caption), "PERCENT")
        End If
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 202074

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 202074
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 202145
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_REPETIDO_NO_GRID", gErr, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202075)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cidade(objGridInt As AdmGrid, Optional ByVal bTestaRepeticao As Boolean = False) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCidade As New ClassCidades
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Cidade

    Set objGridInt.objControle = Cidade

    If Len(Trim(Cidade.Text)) > 0 Then
        
        objCidade.sDescricao = Cidade.Text
        lErro = CF("Cidade_Le_Nome", objCidade)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 202076
        
        If lErro <> SUCESSO Then gError 202077

        If bTestaRepeticao Then
            For iIndice = 1 To objGridInt.iLinhasExistentes
                If iIndice <> objGridInt.objGrid.Row Then
                    If UCase(Cidade.Text) = UCase(objGridInt.objGrid.TextMatrix(iIndice, objGridInt.objGrid.Col)) Then gError 202145
                End If
            Next
        End If
        
        If UCase(Cidade.Text) <> UCase(GridCidades.TextMatrix(GridCidades.Row, iGrid_Cidade_Col)) Then iFiltroAlterado = REGISTRO_ALTERADO
        
        Call Adiciona_Linha(objGridInt)
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 202078

    Saida_Celula_Cidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Cidade:

    Saida_Celula_Cidade = gErr

    Select Case gErr

        Case 202076, 202078
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 202077
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA2", gErr, objCidade.sDescricao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 202145
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_REPETIDO_NO_GRID", gErr, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202079)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_Percentual(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bAdicionaLinha As Boolean = False) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim dPercent As Double

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(objControle.Text)
        If lErro <> SUCESSO Then gError 202080

        dPercent = StrParaDbl(objControle.Text)

        objControle.Text = Format(dPercent, "Fixed")
        
        If bAdicionaLinha Then
            Call Adiciona_Linha(objGridInt)
        End If
        
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, 2) = Round(StrParaInt(QuantidadeTotal.Caption) * (dPercent / 100), 0)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 202081

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = gErr

    Select Case gErr

        Case 202080, 202081
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202082)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_Quantidade(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bAdicionaLinha As Boolean = False) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica o valor informado
        lErro = Valor_Positivo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 202083

        'objControle.Text = Formata_Estoque(objControle.Text)
        
        If bAdicionaLinha Then
            Call Adiciona_Linha(objGridInt)
        End If

        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, 3) = Format(StrParaInt(objControle.Text) / StrParaInt(QuantidadeTotal.Caption), "PERCENT")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 202084

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 202083 To 202084
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202085)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vendedor(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bAdicionaLinha As Boolean = False, Optional ByVal bTestaRepeticao As Boolean = False) As Long
'Faz a crítica da célula Vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim iLinhasAnt As Integer
Dim iIndice As Integer
Dim iQtd As Integer

On Error GoTo Erro_Saida_Celula_Vendedor

    Set objGridInt.objControle = objControle

    'Verifica se vendedor está preenchido
    If Len(Trim(objControle.Text)) > 0 Then

        objControle.Text = LCodigo_Extrai(objControle.Text)
        
        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Le2(objControle, objVendedor, DESMARCADO)
        If lErro <> SUCESSO Then gError 202086
        
        If bTestaRepeticao Then
            For iIndice = 1 To objGridInt.iLinhasExistentes
                If iIndice <> objGridInt.objGrid.Row Then
                    If UCase(objControle.Text) = UCase(objGridInt.objGrid.TextMatrix(iIndice, objGridInt.objGrid.Col)) Then gError 202145
                End If
            Next
        End If
        
        iLinhasAnt = objGridInt.iLinhasExistentes
       
        If bAdicionaLinha Then
            Call Adiciona_Linha(objGridInt)
        End If
        
        If Me.ActiveControl Is VendNovo And iLinhasAnt <> objGridInt.iLinhasExistentes Then
        
            iQtd = 0
            For iIndice = 1 To objGridInt.objGrid.Row - 1
                iQtd = iQtd + StrParaInt(objGridInt.objGrid.TextMatrix(iIndice, 2))
            Next
            
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, 2) = CStr(StrParaInt(QuantidadeTotal.Caption) - iQtd)
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, 3) = Format((StrParaInt(QuantidadeTotal.Caption) - iQtd) / StrParaInt(QuantidadeTotal.Caption), "PERCENT")
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 202087

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = gErr

    Select Case gErr

        Case 202086, 202087
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 202145
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_REPETIDO_NO_GRID", gErr, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202088)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function Adiciona_Linha(ByVal objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Adiciona_Linha
              
    'verifica se precisa preencher o grid com uma nova linha
    If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    Adiciona_Linha = SUCESSO
        
    Exit Function

Erro_Adiciona_Linha:

    Adiciona_Linha = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202089)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 202090

    'Limpa a Tela
    lErro = Limpa_Tela_ClienteExpresso
    If lErro <> SUCESSO Then gError 202091

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 202090, 202091

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202092)

    End Select

End Sub

Function Limpa_Tela_ClienteExpresso() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_ClienteExpresso
        
    gbTrazendoDados = True
    gbLimpandoDados = True
       
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    QuantidadeTotal.Caption = ""
    QuantidadeMatriz.Caption = ""
    
    Call Grid_Limpa(objGridClientes)
    Call Grid_Limpa(objGridCidades)
    Call Grid_Limpa(objGridVendNovo)
    Call Grid_Limpa(objGridCobrNovo)
    Call Grid_Limpa(objGridCallCenterNovo)
    Call Grid_Limpa(objGridBairros)
    
    For iIndice = 0 To UF.ListCount - 1
        UF.Selected(iIndice) = True
    Next
    
    For iIndice = 0 To TiposCliente.ListCount - 1
        TiposCliente.Selected(iIndice) = True
    Next
    
    TransportadoraTodas.Value = vbChecked
    Call TransportadoraTodas_Click
    RegiaoTodas.Value = vbChecked
    Call RegiaoTodas_Click
    VendedorTodos.Value = vbChecked
    Call VendedorTodos_Click
    UsuCobradorTodos.Value = vbChecked
    Call UsuCobradorTodos_Click
    UsuRespCallCenterTodos.Value = vbChecked
    Call UsuRespCallCenterTodos_Click
    CidadesTodas.Value = vbChecked
    Call CidadesTodas_Click
    BairrosTodos.Value = vbChecked
    Call BairrosTodos_Click
    
    Transportadora.ListIndex = -1
    TransportadoraGrid.ListIndex = -1
    TransportadoraNova.ListIndex = -1
    UsuCobrador.ListIndex = -1
    UsuCobradorGrid.ListIndex = -1
    CobrNovo.ListIndex = -1
    UsuRespCallCenter.ListIndex = -1
    UsuRespCallCenterGrid.ListIndex = -1
    CallCenterNovo.ListIndex = -1
    Regiao.ListIndex = -1
    RegiaoGrid.ListIndex = -1
    RegiaoNova.ListIndex = -1
    
    Set gobjClienteExpressoAnt = New ClassClienteExpressoSel
    
    '#####################################
    'Inserido por Wagner
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteDe.Enabled = False
    CategoriaClienteAte.Enabled = False
    CategoriaClienteDe.ListIndex = -1
    CategoriaClienteAte.ListIndex = -1
    '#####################################
      
    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = TAB_Selecao
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    'Torna Frame atual invisível
    Frame2(TabStrip2.SelectedItem.Index).Visible = False
    iFrame2Atual = TAB_Selecao
    'Torna Frame atual visível
    Frame2(iFrame2Atual).Visible = True
    TabStrip2.Tabs.Item(iFrame2Atual).Selected = True
    
    gbLimpandoDados = False
    iAlterado = 0

    Limpa_Tela_ClienteExpresso = SUCESSO

    Exit Function

Erro_Limpa_Tela_ClienteExpresso:

    Limpa_Tela_ClienteExpresso = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202093)

    End Select

    Exit Function

End Function

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 202094
    
    UsuCobrador.Clear
    UsuCobradorGrid.Clear
    CobrNovo.Clear
    UsuRespCallCenter.Clear
    UsuRespCallCenterGrid.Clear
    CallCenterNovo.Clear

    UsuCobrador.AddItem " "
    UsuCobradorGrid.AddItem " "
    CobrNovo.AddItem " "
    UsuRespCallCenter.AddItem " "
    UsuRespCallCenterGrid.AddItem " "
    CallCenterNovo.AddItem " "

    For Each objUsuarios In colUsuarios
        UsuCobrador.AddItem objUsuarios.sCodUsuario
        UsuCobradorGrid.AddItem objUsuarios.sCodUsuario
        CobrNovo.AddItem objUsuarios.sCodUsuario
        UsuRespCallCenter.AddItem objUsuarios.sCodUsuario
        UsuRespCallCenterGrid.AddItem objUsuarios.sCodUsuario
        CallCenterNovo.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case 202094

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202095)

    End Select

    Exit Function

End Function

Private Function Carrega_Regiao() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_Regiao

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 202096
    
    Regiao.Clear
    RegiaoGrid.Clear
    RegiaoNova.Clear

    Regiao.AddItem " "
    Regiao.ItemData(Regiao.NewIndex) = 0

    RegiaoGrid.AddItem " "
    RegiaoGrid.ItemData(RegiaoGrid.NewIndex) = 0

    RegiaoNova.AddItem " "
    RegiaoNova.ItemData(RegiaoNova.NewIndex) = 0

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Regiao.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoDescricao.iCodigo
    
        RegiaoGrid.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        RegiaoGrid.ItemData(RegiaoGrid.NewIndex) = objCodigoDescricao.iCodigo
    
        RegiaoNova.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        RegiaoNova.ItemData(RegiaoNova.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Carrega_Regiao = SUCESSO

    Exit Function

Erro_Carrega_Regiao:

    Carrega_Regiao = gErr

    Select Case gErr

        Case 202096

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202097)

    End Select

    Exit Function

End Function

Private Function Carrega_Transportadora() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 202098
    
    Transportadora.Clear
    TransportadoraGrid.Clear
    TransportadoraNova.Clear

    Transportadora.AddItem " "
    Transportadora.ItemData(Transportadora.NewIndex) = 0

    TransportadoraGrid.AddItem " "
    TransportadoraGrid.ItemData(TransportadoraGrid.NewIndex) = 0

    TransportadoraNova.AddItem " "
    TransportadoraNova.ItemData(TransportadoraNova.NewIndex) = 0

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Transportadora.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoDescricao.iCodigo
    
        TransportadoraGrid.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TransportadoraGrid.ItemData(TransportadoraGrid.NewIndex) = objCodigoDescricao.iCodigo
    
        TransportadoraNova.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TransportadoraNova.ItemData(TransportadoraNova.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Carrega_Transportadora = SUCESSO

    Exit Function

Erro_Carrega_Transportadora:

    Carrega_Transportadora = gErr

    Select Case gErr

        Case 202098

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202099)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim colClientes As New Collection
Dim colFiliais As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Move_Tela_Memoria(colClientes, colFiliais)
    If lErro <> SUCESSO Then gError 202100
    
    If colClientes.Count = 0 And colFiliais.Count = 0 Then gError 202101
    
    lErro = CF("ClienteExpresso_Grava", colClientes, colFiliais)
    If lErro <> SUCESSO Then gError 202102
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 202100, 202102
        
        Case 202101
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_CLIENTE_ALTERADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202103)
        
    End Select
    
End Function

Public Function Move_Tela_Memoria(ByVal colClientes As Collection, ByVal colFiliais As Collection) As Long

Dim lErro As Long
Dim objcliente As ClassCliente
Dim objFilial As ClassFilialCliente
Dim objClienteAux As ClassCliente
Dim objClienteCol As ClassCliente
Dim objFilialAux As ClassFilialCliente
Dim iLinha As Integer

On Error GoTo Erro_Move_Tela_Memoria

    For iLinha = 1 To objGridClientes.iLinhasExistentes
    
        Set objClienteAux = New ClassCliente
        Set objFilialAux = New ClassFilialCliente
        Set objcliente = gobjClienteExpressoAnt.colClientes.Item(iLinha)
        Set objFilial = gobjClienteExpressoAnt.colFiliais.Item(iLinha)
        
        objClienteAux.lCodigo = LCodigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_Cliente_Col))
        objFilialAux.lCodCliente = objClienteAux.lCodigo
        objFilialAux.iCodFilial = Codigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_FilialCliente_Col))
        objFilialAux.iRegiao = Codigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_Regiao_Col))
        objFilialAux.iVendedor = Codigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_Vendedor_Col))
        objFilialAux.iCodTransportadora = Codigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_Transp_Col))
        objClienteAux.sUsuarioCobrador = GridClientes.TextMatrix(iLinha, iGrid_Cobr_Col)
        objClienteAux.sUsuRespCallCenter = GridClientes.TextMatrix(iLinha, iGrid_CallCenter_Col)
        
        'Se trocou alguma informação coloca na oleção para gravar
        If objFilialAux.iRegiao <> objFilial.iRegiao Or _
            objFilialAux.iVendedor <> objFilial.iVendedor Or _
            objFilialAux.iCodTransportadora <> objFilial.iCodTransportadora Then
            
            colFiliais.Add objFilialAux
        End If
        
        'Se trocou alguma informação coloca na oleção para gravar
        If objClienteAux.sUsuarioCobrador <> objcliente.sUsuarioCobrador Or _
            objClienteAux.sUsuRespCallCenter <> objcliente.sUsuRespCallCenter Then
            
            'Só altera o cliente para a filial matriz
            If objFilialAux.iCodFilial = FILIAL_MATRIZ Then colClientes.Add objClienteAux
        End If
    
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202104)
        
    End Select
    
End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Private Sub RegiaoGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RegiaoGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Private Sub RegiaoGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Private Sub RegiaoGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = RegiaoGrid
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VendedorGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VendedorGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Private Sub VendedorGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Private Sub VendedorGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = VendedorGrid
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TransportadoraGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TransportadoraGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Private Sub TransportadoraGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Private Sub TransportadoraGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = TransportadoraGrid
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UsuCobradorGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UsuCobradorGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Private Sub UsuCobradorGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Private Sub UsuCobradorGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = UsuCobradorGrid
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UsuRespCallCenterGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UsuRespCallCenterGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Private Sub UsuRespCallCenterGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Private Sub UsuRespCallCenterGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = UsuRespCallCenterGrid
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CobrNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CobrNovo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCobrNovo)
End Sub

Private Sub CobrNovo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCobrNovo)
End Sub

Private Sub CobrNovo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCobrNovo.objControle = CobrNovo
    lErro = Grid_Campo_Libera_Foco(objGridCobrNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CobrNovoQtd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CobrNovoQtd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCobrNovo)
End Sub

Private Sub CobrNovoQtd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCobrNovo)
End Sub

Private Sub CobrNovoQtd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCobrNovo.objControle = CobrNovoQtd
    lErro = Grid_Campo_Libera_Foco(objGridCobrNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CobrNovoPerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CobrNovoPerc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCobrNovo)
End Sub

Private Sub CobrNovoPerc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCobrNovo)
End Sub

Private Sub CobrNovoPerc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCobrNovo.objControle = CobrNovoPerc
    lErro = Grid_Campo_Libera_Foco(objGridCobrNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CallCenterNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CallCenterNovo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCallCenterNovo)
End Sub

Private Sub CallCenterNovo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCallCenterNovo)
End Sub

Private Sub CallCenterNovo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCallCenterNovo.objControle = CallCenterNovo
    lErro = Grid_Campo_Libera_Foco(objGridCallCenterNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CallCenterNovoQtd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CallCenterNovoQtd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCallCenterNovo)
End Sub

Private Sub CallCenterNovoQtd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCallCenterNovo)
End Sub

Private Sub CallCenterNovoQtd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCallCenterNovo.objControle = CallCenterNovoQtd
    lErro = Grid_Campo_Libera_Foco(objGridCallCenterNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CallCenterNovoPerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CallCenterNovoPerc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCallCenterNovo)
End Sub

Private Sub CallCenterNovoPerc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCallCenterNovo)
End Sub

Private Sub CallCenterNovoPerc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCallCenterNovo.objControle = CallCenterNovoPerc
    lErro = Grid_Campo_Libera_Foco(objGridCallCenterNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VendNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VendNovo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVendNovo)
End Sub

Private Sub VendNovo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVendNovo)
End Sub

Private Sub VendNovo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridVendNovo.objControle = VendNovo
    lErro = Grid_Campo_Libera_Foco(objGridVendNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VendNovoQtd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VendNovoQtd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVendNovo)
End Sub

Private Sub VendNovoQtd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVendNovo)
End Sub

Private Sub VendNovoQtd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridVendNovo.objControle = VendNovoQtd
    lErro = Grid_Campo_Libera_Foco(objGridVendNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VendNovoPerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VendNovoPerc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVendNovo)
End Sub

Private Sub VendNovoPerc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVendNovo)
End Sub

Private Sub VendNovoPerc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridVendNovo.objControle = VendNovoPerc
    lErro = Grid_Campo_Libera_Foco(objGridVendNovo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Cidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCidades)
End Sub

Private Sub Cidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCidades)
End Sub

Private Sub Cidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCidades.objControle = Cidade
    lErro = Grid_Campo_Libera_Foco(objGridCidades)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Bairro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Bairro_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBairros)
End Sub

Private Sub Bairro_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBairros)
End Sub

Private Sub Bairro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBairros.objControle = Bairro
    lErro = Grid_Campo_Libera_Foco(objGridBairros)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CategoriaCliente_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaCliente_Click

    If Len(Trim(CategoriaCliente.Text)) > 0 Then
        CategoriaClienteDe.Enabled = True
        CategoriaClienteAte.Enabled = True
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteDe)
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteAte)
    Else
        CategoriaClienteDe.Enabled = False
        CategoriaClienteAte.Enabled = False
    End If

    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202105)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ComboCategoriaCliente(ByVal objCombo As ComboBox) As Long

Dim lErro As Long
Dim colCategoriaCliente As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then gError 202106

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente
        objCombo.AddItem objCategoriaCliente.sCategoria
    Next
    
    Carrega_ComboCategoriaCliente = SUCESSO
    
    Exit Function

Erro_Carrega_ComboCategoriaCliente:

    Carrega_ComboCategoriaCliente = gErr

    Select Case gErr
    
        Case 202106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202107)

    End Select

    Exit Function

End Function

Private Sub Carrega_ComboCategoriaItens(ByVal objComboCategoria As ComboBox, ByVal objComboItens As ComboBox)

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_Carrega_ComboCategoriaItens

    'Verifica se a CategoriaCliente foi preenchida
    If objComboCategoria.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = objComboCategoria.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then gError 202108

        objComboItens.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        objComboItens.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria
            objComboItens.AddItem objCategoriaClienteItem.sItem
        Next
        
        CategoriaClienteTodas.Value = vbFalse
    
    Else
        
        'Senão Desablita ItemCategoriaCliente
        objComboItens.ListIndex = -1
        objComboItens.Enabled = False
    
    End If
    
    Exit Sub

Erro_Carrega_ComboCategoriaItens:

    Select Case gErr
    
        Case 202108

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202109)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 202110
        
        If lErro <> SUCESSO Then gError 202111
    
    End If
    
    'Se a CategoriaCliente estiver em branco desabilita e limpa a combo
    If Len(CategoriaCliente.Text) = 0 Then
        CategoriaClienteDe.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Enabled = False
        CategoriaClienteAte.Clear
    End If
    
    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 202110
         
        Case 202111
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaCliente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202112)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteItem_Validate(Cancel As Boolean, objCombo As ComboBox)

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteItem_Validate

    If Len(objCombo.Text) <> 0 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(objCombo)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 202113
        
        If lErro <> SUCESSO Then gError 202114
    
    End If

    Exit Sub

Erro_CategoriaClienteItem_Validate:

    Cancel = True

    Select Case gErr

        Case 202113
        
        Case 202114
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", gErr, objCombo.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202115)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteTodas_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteTodas_Click

    If CategoriaClienteTodas.Value = vbChecked Then
        'Desabilita o combotipo
        CategoriaCliente.ListIndex = -1
        CategoriaCliente.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Clear
    Else
        CategoriaCliente.Enabled = True
    End If

    Call CategoriaCliente_Click

    Exit Sub

Erro_CategoriaClienteTodas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202116)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteAte_Validate(Cancel As Boolean)
    Call CategoriaClienteItem_Validate(Cancel, CategoriaClienteAte)
End Sub

Private Sub CategoriaClienteDe_Validate(Cancel As Boolean)
    Call CategoriaClienteItem_Validate(Cancel, CategoriaClienteDe)
End Sub

Private Sub LabelClienteAte_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    If Len(Trim(ClienteAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objcliente.lCodigo = LCodigo_Extrai(ClienteAte.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 1

    If Len(Trim(ClienteDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objcliente.lCodigo = LCodigo_Extrai(ClienteDe.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente

    Set objcliente = obj1
    
    'Preenche campo Cliente
    If giClienteInicial = 1 Then
        ClienteDe.Text = CStr(objcliente.lCodigo)
        Call ClienteDe_Validate(bSGECancelDummy)
    Else
        ClienteAte.Text = CStr(objcliente.lCodigo)
        Call ClienteAte_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is Cidade Then
            Call BotaoCidades_Click
        ElseIf Me.ActiveControl Is VendedorGrid Then
            Call BotaoVendedorGrid_Click
        ElseIf Me.ActiveControl Is VendNovo Then
            Call BotaoVendedor_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call VendedorLabel_Click
        ElseIf Me.ActiveControl Is Bairro Then
            Call BotaoBairros_Click
        End If
          
    End If

End Sub

Private Sub ClienteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_ClienteDe_Validate

    If Len(Trim(ClienteDe.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteDe, objcliente, 0)
        If lErro <> SUCESSO Then gError 202117

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 202117
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202118)

    End Select

End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_ClienteAte_Validate

    If Len(Trim(ClienteAte.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteAte, objcliente, 0)
        If lErro <> SUCESSO Then gError 202119

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 202119
             Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objcliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202120)

    End Select

End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor)
        If lErro <> SUCESSO Then gError 202121

    End If
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 202121
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202122)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim iIndice As Integer
Dim iLinhasAnt As Integer
Dim iQtd As Integer

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1

    If giVendedor = VENDEDOR_FIL Then
        Vendedor.Text = objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido
    ElseIf giVendedor = VENDEDOR_CLI Then
    
        If Me.ActiveControl Is VendedorGrid Then
            VendedorGrid.Text = objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido
        Else
            GridClientes.TextMatrix(GridClientes.Row, iGrid_Vendedor_Col) = objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido
        End If
    
    Else
    
        For iIndice = 1 To objGridVendNovo.iLinhasExistentes
            If iIndice <> GridVendNovo.Row Then
                If objVendedor.iCodigo = Codigo_Extrai(GridVendNovo.TextMatrix(iIndice, iGrid_VendNovo_Col)) Then gError 202145
            End If
        Next
    
        If Me.ActiveControl Is VendNovo Then
            VendNovo.Text = objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido
        Else
            GridVendNovo.TextMatrix(GridVendNovo.Row, iGrid_VendNovo_Col) = objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido
        End If
        
        iLinhasAnt = objGridVendNovo.iLinhasExistentes
        
        Call Adiciona_Linha(objGridVendNovo)
        
        If iLinhasAnt <> objGridVendNovo.iLinhasExistentes Then
        
            iQtd = 0
            For iIndice = 1 To GridVendNovo.Row - 1
                iQtd = iQtd + StrParaInt(GridVendNovo.TextMatrix(iIndice, iGrid_VendNovoQtd_Col))
            Next
            
            GridVendNovo.TextMatrix(GridVendNovo.Row, iGrid_VendNovoQtd_Col) = CStr(StrParaInt(QuantidadeTotal.Caption) - iQtd)
            GridVendNovo.TextMatrix(GridVendNovo.Row, iGrid_VendNovoPerc_Col) = Format((StrParaInt(QuantidadeTotal.Caption) - iQtd) / StrParaInt(QuantidadeTotal.Caption), "PERCENT")
        End If
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr
        
        Case 202145
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_REPETIDO_NO_GRID", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202123)
    
    End Select
    
    Exit Sub

End Sub

Public Sub VendedorLabel_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection

On Error GoTo Erro_VendedorLabel_Click

    If Not Vendedor.Enabled Then Exit Sub
    
    giVendedor = VENDEDOR_FIL
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

    Exit Sub

Erro_VendedorLabel_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202124)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoVendedorGrid_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVendedorGrid_Click
    
    giVendedor = VENDEDOR_CLI
    
    If GridClientes.Row = 0 Then gError 202125
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Me.ActiveControl Is VendNovo Then
        objVendedor.iCodigo = Codigo_Extrai(VendedorGrid.Text)
        objVendedor.sNomeReduzido = VendedorGrid.Text
    Else
        objVendedor.iCodigo = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Vendedor_Col))
    End If
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

    Exit Sub

Erro_BotaoVendedorGrid_Click:

    Select Case gErr
        
        Case 202125
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202126)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVendedor_Click
    
    giVendedor = VENDEDOR_NOV
    
    If GridVendNovo.Row = 0 Then gError 202127
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Me.ActiveControl Is VendNovo Then
        objVendedor.iCodigo = Codigo_Extrai(VendNovo.Text)
        objVendedor.sNomeReduzido = VendNovo.Text
    Else
        objVendedor.iCodigo = Codigo_Extrai(GridVendNovo.TextMatrix(GridVendNovo.Row, iGrid_VendNovo_Col))
    End If
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

    Exit Sub

Erro_BotaoVendedor_Click:

    Select Case gErr
        
        Case 202127
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202128)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades
Dim iIndice As Integer

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1
    
    For iIndice = 1 To objGridCidades.iLinhasExistentes
        If iIndice <> GridCidades.Row Then
            If UCase(objCidade.sDescricao) = UCase(GridCidades.TextMatrix(iIndice, iGrid_Cidade_Col)) Then gError 202145
        End If
    Next
    
    If Me.ActiveControl Is Cidade Then
        Cidade.Text = objCidade.sDescricao
    Else
        GridCidades.TextMatrix(GridCidades.Row, iGrid_Cidade_Col) = objCidade.sDescricao
    End If
    
    Call Adiciona_Linha(objGridCidades)

    Me.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr

        Case 202145
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_REPETIDO_NO_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202129)

    End Select

    Exit Sub

End Sub

Public Sub BotaoCidades_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

On Error GoTo Erro_BotaoCidades_Click

    If GridCidades.Row = 0 Then gError 202130

    If Me.ActiveControl Is Cidade Then
        objCidade.sDescricao = Cidade.Text
    Else
        objCidade.sDescricao = GridCidades.TextMatrix(GridCidades.Row, iGrid_Cidade_Col)
    End If
    
    'Chama a Tela de browse
    Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)

    Exit Sub

Erro_BotaoCidades_Click:

    Select Case gErr

        Case 202130
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202131)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCliente_Click()

Dim objcliente As New ClassCliente

On Error GoTo Erro_BotaoCliente_Click

    If GridClientes.Row = 0 Then gError 202132
       
    objcliente.lCodigo = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))

    Call Chama_Tela("Clientes", objcliente)

    Exit Sub

Erro_BotaoCliente_Click:

    Select Case gErr
    
        Case 202132
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202133)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHistoricoCR_Click()

Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHistoricoCR_Click

    If GridClientes.Row = 0 Then gError 202134
       
    colSelecao.Add LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))

    Call Chama_Tela("TitRecTodosTFLista", colSelecao, Nothing, Nothing, "Cliente = ?")

    Exit Sub

Erro_BotaoHistoricoCR_Click:

    Select Case gErr
    
        Case 202134
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202135)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHistoricoCRM_Click()

Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHistoricoCRM_Click

    If GridClientes.Row = 0 Then gError 202136
       
    colSelecao.Add LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))

    Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, Nothing, Nothing, "ClienteCod = ?")

    Exit Sub

Erro_BotaoHistoricoCRM_Click:

    Select Case gErr
    
        Case 202136
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202137)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAplicar_Click(Index As Integer)

Dim iIndice As Integer
Dim iCount As Integer
Dim aValores() As Variant
Dim colClientes  As New Collection
Dim objcliente As ClassCliente
Dim iQtd As Integer
Dim iLinha As Integer
Dim bSoMatriz As Boolean
Dim bAchou As Boolean

On Error GoTo Erro_BotaoAplicar_Click

    iCount = 1
    bSoMatriz = False

    Select Case Index
    
        Case FRAME2_REGIAO
            
            ReDim aValores(1 To iCount, 1 To 2)
            
            aValores(1, 1) = RegiaoNova.Text
            aValores(1, 2) = StrParaInt(QuantidadeTotal.Caption)
    
        Case FRAME2_VEND
        
            iCount = objGridVendNovo.iLinhasExistentes
            If iCount = 0 Then iCount = 1
            ReDim aValores(1 To iCount, 1 To 2)
        
            For iIndice = 1 To iCount
                aValores(iIndice, 1) = GridVendNovo.TextMatrix(iIndice, iGrid_VendNovo_Col)
                aValores(iIndice, 2) = StrParaInt(GridVendNovo.TextMatrix(iIndice, iGrid_VendNovoQtd_Col))
            Next
    
        Case FRAME2_TRANSP
            
            ReDim aValores(1 To iCount, 1 To 2)
            
            aValores(1, 1) = TransportadoraNova.Text
            aValores(1, 2) = StrParaInt(QuantidadeTotal.Caption)
    
        Case FRAME2_COBR
        
            iCount = objGridCobrNovo.iLinhasExistentes
            If iCount = 0 Then iCount = 1
            ReDim aValores(1 To iCount, 1 To 2)
        
            For iIndice = 1 To iCount
                aValores(iIndice, 1) = GridCobrNovo.TextMatrix(iIndice, iGrid_CobrNovo_Col)
                aValores(iIndice, 2) = StrParaInt(GridCobrNovo.TextMatrix(iIndice, iGrid_CobrNovoQtd_Col))
            Next
            
            bSoMatriz = True
            
        Case FRAME2_CALLCENTER
        
            iCount = objGridCallCenterNovo.iLinhasExistentes
            If iCount = 0 Then iCount = 1
            ReDim aValores(1 To iCount, 1 To 2)
        
            For iIndice = 1 To iCount
                aValores(iIndice, 1) = GridCallCenterNovo.TextMatrix(iIndice, iGrid_CallCenterNovo_Col)
                aValores(iIndice, 2) = StrParaInt(GridCallCenterNovo.TextMatrix(iIndice, iGrid_CallCenterNovoQtd_Col))
            Next
            
            bSoMatriz = True
    
    End Select
    
    iQtd = 0
    For iIndice = 1 To iCount
        iQtd = iQtd + aValores(iIndice, 2)
    Next
    
    If Not bSoMatriz Then
        If iQtd <> StrParaInt(QuantidadeTotal.Caption) Then gError 202138
    Else
        If iQtd <> StrParaInt(QuantidadeMatriz.Caption) Then gError 202139
    End If
    
    iLinha = 0
    For iIndice = 1 To iCount
        For iQtd = 1 To aValores(iIndice, 2)
            iLinha = iLinha + 1
            If Not bSoMatriz Or Codigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_FilialCliente_Col)) = FILIAL_MATRIZ Then
                GridClientes.TextMatrix(iLinha, Index + 5) = aValores(iIndice, 1)
                If bSoMatriz Then
                    Set objcliente = New ClassCliente
                    objcliente.lCodigo = LCodigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_Cliente_Col))
                    objcliente.sGuia = aValores(iIndice, 1)
                    colClientes.Add objcliente
                End If
            End If
        Next
    Next

    If bSoMatriz Then
        For iLinha = 1 To StrParaInt(QuantidadeTotal.Caption)
            bAchou = False
            If Codigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_FilialCliente_Col)) <> FILIAL_MATRIZ Then
                For Each objcliente In colClientes
                    If objcliente.lCodigo = LCodigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_Cliente_Col)) Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If bAchou Then
                    GridClientes.TextMatrix(iLinha, Index + 5) = objcliente.sGuia
                End If
            End If
        Next
    End If
    
    Exit Sub

Erro_BotaoAplicar_Click:

    Select Case gErr
    
        Case 202138
             Call Rotina_Erro(vbOKOnly, "ERRO_QTD_TOTAL_DIFERENTE", gErr, CStr(iQtd), QuantidadeTotal.Caption)
    
        Case 202139
             Call Rotina_Erro(vbOKOnly, "ERRO_QTD_TOTAL_DIFERENTE", gErr, CStr(iQtd), QuantidadeMatriz.Caption)
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202140)

    End Select

    Exit Sub
    
End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
            
        Case Bairro.Name, Cidade.Name, VendedorGrid.Name, RegiaoGrid.Name, TransportadoraGrid.Name, VendNovo.Name, CobrNovo.Name, CallCenterNovo.Name
            objControl.Enabled = True
            
        Case UsuCobradorGrid.Name, UsuRespCallCenterGrid.Name
            If Codigo_Extrai(GridClientes.TextMatrix(iLinha, iGrid_FilialCliente_Col)) = FILIAL_MATRIZ Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case VendNovoQtd.Name, VendNovoPerc.Name
        
            If Len(Trim(GridVendNovo.TextMatrix(iLinha, iGrid_VendNovo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case CobrNovoQtd.Name, CobrNovoPerc.Name
        
            If Len(Trim(GridCobrNovo.TextMatrix(iLinha, iGrid_CobrNovo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case CallCenterNovoQtd.Name, CallCenterNovoPerc.Name
        
            If Len(Trim(GridCallCenterNovo.TextMatrix(iLinha, iGrid_CallCenterNovo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case ClienteGrid.Name, FilialClienteGrid.Name, CidadeGrid.Name, UFGrid.Name, BairroGrid.Name
            objControl.Enabled = False
     
        Case Else
            objControl.Enabled = False
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202141)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 202142

    'Limpa a tela
    Call Limpa_Tela_ClienteExpresso

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 202142

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202143)

    End Select
    
    Exit Sub

End Sub

Private Sub CidadesTodas_Click()
    If CidadesTodas.Value = vbChecked Then
        FrameCidade.Enabled = False
        Call Grid_Limpa(objGridCidades)
    Else
        FrameCidade.Enabled = True
    End If
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BairrosTodos_Click()
    If BairrosTodos.Value = vbChecked Then
        FrameBairro.Enabled = False
        Call Grid_Limpa(objGridBairros)
    Else
        FrameBairro.Enabled = True
    End If
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RegiaoTodas_Click()
    If RegiaoTodas.Value = vbChecked Then
        Regiao.Enabled = False
        Regiao.ListIndex = -1
    Else
        Regiao.Enabled = True
    End If
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TransportadoraTodas_Click()
    If TransportadoraTodas.Value = vbChecked Then
        Transportadora.Enabled = False
        Transportadora.ListIndex = -1
    Else
        Transportadora.Enabled = True
    End If
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VendedorTodos_Click()
    If VendedorTodos.Value = vbChecked Then
        Vendedor.Enabled = False
        VendedorLabel.MousePointer = vbDefault
        Vendedor.Text = ""
    Else
        Vendedor.Enabled = True
        VendedorLabel.MousePointer = vbArrowQuestion
    End If
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UsuCobradorTodos_Click()
    If UsuCobradorTodos.Value = vbChecked Then
        UsuCobrador.Enabled = False
        UsuCobrador.ListIndex = -1
    Else
        UsuCobrador.Enabled = True
    End If
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UsuRespCallCenterTodos_Click()
    If UsuRespCallCenterTodos.Value = vbChecked Then
        UsuRespCallCenter.Enabled = False
        UsuRespCallCenter.ListIndex = -1
    Else
        UsuRespCallCenter.Enabled = True
    End If
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_Estados() As Long

Dim lErro As Long
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Estados

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError 202098
    
    UF.Clear

    iIndice = -1
    For Each vCodigo In colCodigo
        iIndice = iIndice + 1
        UF.AddItem vCodigo
        UF.Selected(iIndice) = True
    Next

    Carrega_Estados = SUCESSO

    Exit Function

Erro_Carrega_Estados:

    Carrega_Estados = gErr

    Select Case gErr

        Case 202098

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202099)

    End Select

    Exit Function

End Function

Private Sub objEventoBairro_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objBairro As ClassEndereco
Dim iIndice As Integer

On Error GoTo Erro_objEventoBairro_evSelecao

    Set objBairro = obj1

    For iIndice = 1 To objGridBairros.iLinhasExistentes
        If iIndice <> GridBairros.Row Then
            If UCase(objBairro.sBairro) = UCase(GridBairros.TextMatrix(iIndice, iGrid_Bairro_Col)) Then gError 202145
        End If
    Next
            
    If Me.ActiveControl Is Bairro Then
        Bairro.Text = objBairro.sBairro
    Else
        GridBairros.TextMatrix(GridBairros.Row, iGrid_Bairro_Col) = objBairro.sBairro
    End If

    Call Adiciona_Linha(objGridBairros)

    Me.Show

    Exit Sub

Erro_objEventoBairro_evSelecao:

    Select Case gErr

        Case 202145
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_REPETIDO_NO_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202129)

    End Select

    Exit Sub

End Sub

Public Sub BotaoBairros_Click()

Dim objBairro As New ClassEndereco
Dim colSelecao As Collection

On Error GoTo Erro_BotaoBairros_Click

    If GridBairros.Row = 0 Then gError 202130

    If Me.ActiveControl Is Bairro Then
        objBairro.sBairro = Bairro.Text
    Else
        objBairro.sBairro = GridBairros.TextMatrix(GridBairros.Row, iGrid_Bairro_Col)
    End If
    
    'Chama a Tela de browse
    Call Chama_Tela("BairrosLista", colSelecao, objBairro, objEventoBairro)

    Exit Sub

Erro_BotaoBairros_Click:

    Select Case gErr

        Case 202130
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202131)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoDesmarcarTodos_Click(Index As Integer)
Dim iLinha As Integer
    iFiltroAlterado = REGISTRO_ALTERADO
    
    If Index <> 0 Then
        For iLinha = 0 To UF.ListCount - 1
            UF.Selected(iLinha) = False
        Next
    Else
        For iLinha = 0 To TiposCliente.ListCount - 1
            TiposCliente.Selected(iLinha) = False
        Next
    End If
End Sub

Private Sub BotaoMarcarTodos_Click(Index As Integer)
Dim iLinha As Integer
    iFiltroAlterado = REGISTRO_ALTERADO

    If Index <> 0 Then
        For iLinha = 0 To UF.ListCount - 1
            UF.Selected(iLinha) = True
        Next
    Else
        For iLinha = 0 To TiposCliente.ListCount - 1
            TiposCliente.Selected(iLinha) = True
        Next
    End If
End Sub

Private Sub BotaoLimparTodos_Click(Index As Integer)
Dim iLinha As Integer
    For iLinha = 1 To objGridClientes.iLinhasExistentes
        GridClientes.TextMatrix(iLinha, Index) = ""
    Next
End Sub

Private Sub BotaoLimparGrid_Click(Index As Integer)
    Select Case Index
        Case FRAME2_VEND
            Call Grid_Limpa(objGridVendNovo)
        Case FRAME2_COBR
            Call Grid_Limpa(objGridCobrNovo)
        Case FRAME2_CALLCENTER
            Call Grid_Limpa(objGridCallCenterNovo)
    End Select
End Sub

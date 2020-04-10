VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ConciliacaoBancariaOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5370
      Index           =   2
      Left            =   210
      TabIndex        =   8
      Top             =   495
      Visible         =   0   'False
      Width           =   9120
      Begin VB.CommandButton Botao_Conciliar 
         Caption         =   "Conciliar"
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
         Left            =   60
         TabIndex        =   18
         Top             =   2490
         Width           =   1020
      End
      Begin VB.Frame Frame5 
         Caption         =   "Procurar Por"
         Height          =   705
         Left            =   4500
         TabIndex        =   45
         Top             =   2340
         Width           =   4545
         Begin VB.CommandButton Botao_ProcurarExt 
            Height          =   330
            Left            =   3570
            Picture         =   "ConciliacaoBancariaOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   255
            Width           =   375
         End
         Begin VB.CommandButton Botao_ProcurarMov 
            Height          =   330
            Left            =   4065
            Picture         =   "ConciliacaoBancariaOcx.ctx":01C2
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   255
            Width           =   375
         End
         Begin VB.ComboBox Procura 
            Height          =   315
            ItemData        =   "ConciliacaoBancariaOcx.ctx":0384
            Left            =   105
            List            =   "ConciliacaoBancariaOcx.ctx":0391
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   255
            Width           =   3390
         End
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "ConciliacaoBancariaOcx.ctx":03AD
         Left            =   1380
         List            =   "ConciliacaoBancariaOcx.ctx":03BA
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   75
         Width           =   2925
      End
      Begin VB.CommandButton Botao_Desconciliar 
         Caption         =   "Desconciliar"
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
         Left            =   1140
         TabIndex        =   19
         Top             =   2490
         Width           =   1245
      End
      Begin VB.Frame Frame4 
         Caption         =   "Exibir Correspondentes"
         Height          =   705
         Left            =   2520
         TabIndex        =   48
         Top             =   2340
         Width           =   1890
         Begin VB.CommandButton Botao_ExibirMov 
            Height          =   345
            Left            =   1020
            Picture         =   "ConciliacaoBancariaOcx.ctx":03D6
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   255
            Width           =   360
         End
         Begin VB.CommandButton Botao_ExibirExt 
            Height          =   345
            Left            =   405
            Picture         =   "ConciliacaoBancariaOcx.ctx":0598
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   255
            Width           =   360
         End
      End
      Begin VB.TextBox HistoricoMov 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3255
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   3660
         Width           =   3525
      End
      Begin VB.CheckBox ConciliadoMov 
         Enabled         =   0   'False
         Height          =   240
         Left            =   750
         TabIndex        =   26
         Top             =   3960
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CheckBox SelecionadoMov 
         Height          =   240
         Left            =   150
         TabIndex        =   25
         Top             =   3615
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Historico 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3000
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1095
         Width           =   3525
      End
      Begin VB.CheckBox Conciliado 
         Enabled         =   0   'False
         Height          =   240
         Left            =   1260
         TabIndex        =   11
         Top             =   585
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox Categoria 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   6240
         TabIndex        =   15
         Top             =   600
         Width           =   450
      End
      Begin VB.TextBox CodLctoBanco 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   6705
         TabIndex        =   16
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox Documento 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   7260
         TabIndex        =   50
         Top             =   600
         Width           =   1470
      End
      Begin VB.CheckBox Selecionado 
         Height          =   240
         Left            =   510
         TabIndex        =   10
         Top             =   570
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   240
         Left            =   2010
         TabIndex        =   12
         Top             =   585
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   240
         Left            =   3165
         TabIndex        =   13
         Top             =   615
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NumRefExterna 
         Height          =   240
         Left            =   6945
         TabIndex        =   30
         Top             =   3660
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataMov 
         Height          =   240
         Left            =   1200
         TabIndex        =   27
         Top             =   3660
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ValorMov 
         Height          =   240
         Left            =   2265
         TabIndex        =   28
         Top             =   3630
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridExtrato 
         Height          =   1755
         Left            =   -30
         TabIndex        =   17
         Top             =   645
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   3096
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid GridMov 
         Height          =   1800
         Left            =   -30
         TabIndex        =   31
         Top             =   3195
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   3175
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label Total1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5100
         TabIndex        =   69
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label Label13 
         Caption         =   "Extrato:"
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
         Left            =   4395
         TabIndex        =   68
         Top             =   135
         Width           =   690
      End
      Begin VB.Label Total3 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7650
         TabIndex        =   67
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label Label11 
         Caption         =   "Sistema:"
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
         Left            =   6900
         TabIndex        =   66
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label7 
         Caption         =   "Movimentos no Extrato Bancário:"
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
         TabIndex        =   61
         Top             =   435
         Width           =   4260
      End
      Begin VB.Label Label4 
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
         Height          =   240
         Left            =   45
         TabIndex        =   62
         Top             =   105
         Width           =   1410
      End
      Begin VB.Label Label5 
         Caption         =   "Movimentos no Sistema:"
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
         Left            =   0
         TabIndex        =   63
         Top             =   2985
         Width           =   4260
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4950
      Index           =   1
      Left            =   225
      TabIndex        =   0
      Top             =   600
      Width           =   9060
      Begin VB.ComboBox CodCCI 
         Height          =   315
         Left            =   1965
         TabIndex        =   1
         Top             =   435
         Width           =   2595
      End
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Movimentos"
         Height          =   3750
         Left            =   1185
         TabIndex        =   46
         Top             =   1065
         Width           =   6270
         Begin VB.OptionButton NaoConciliados 
            Caption         =   "Não Conciliados"
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
            Left            =   1005
            TabIndex        =   6
            Top             =   2970
            Value           =   -1  'True
            Width           =   1770
         End
         Begin VB.OptionButton Todos 
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
            Height          =   255
            Left            =   3570
            TabIndex        =   7
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "Periodo"
            Height          =   945
            Left            =   360
            TabIndex        =   47
            Top             =   495
            Width           =   5505
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   780
               TabIndex        =   2
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   300
               Left            =   1935
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   3465
               TabIndex        =   3
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDown2 
               Height          =   300
               Left            =   4590
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label3 
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
               Height          =   195
               Left            =   2985
               TabIndex        =   54
               Top             =   420
               Width           =   360
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Left            =   345
               TabIndex        =   55
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Valores"
            Height          =   855
            Left            =   360
            TabIndex        =   52
            Top             =   1680
            Width           =   5520
            Begin MSMask.MaskEdBox ValorDe 
               Height          =   300
               Left            =   750
               TabIndex        =   4
               Top             =   345
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorAte 
               Height          =   300
               Left            =   3465
               TabIndex        =   5
               Top             =   345
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label14 
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
               Left            =   315
               TabIndex        =   56
               Top             =   405
               Width           =   315
            End
            Begin VB.Label Label6 
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
               Height          =   195
               Left            =   2985
               TabIndex        =   57
               Top             =   375
               Width           =   360
            End
         End
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   1080
         TabIndex        =   58
         Top             =   465
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5325
      Index           =   3
      Left            =   195
      TabIndex        =   32
      Top             =   540
      Visible         =   0   'False
      Width           =   9165
      Begin VB.CommandButton BotaoDocOriginal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   6195
         Picture         =   "ConciliacaoBancariaOcx.ctx":075A
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4335
         Width           =   2055
      End
      Begin VB.ComboBox Ordenados1 
         Height          =   315
         ItemData        =   "ConciliacaoBancariaOcx.ctx":3670
         Left            =   1860
         List            =   "ConciliacaoBancariaOcx.ctx":367D
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   150
         Width           =   3135
      End
      Begin VB.TextBox HistoricoMov1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3135
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   540
         Width           =   3690
      End
      Begin VB.CheckBox ConciliadoMov1 
         Enabled         =   0   'False
         Height          =   240
         Left            =   900
         TabIndex        =   35
         Top             =   1605
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton Botao_ConciliarMov1 
         Caption         =   "Conciliar"
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
         Left            =   3075
         TabIndex        =   41
         Top             =   4290
         Width           =   1020
      End
      Begin VB.CommandButton Botao_DesconciliarMov1 
         Caption         =   "Desconciliar"
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
         Left            =   4440
         TabIndex        =   42
         Top             =   4290
         Width           =   1245
      End
      Begin VB.CheckBox SelecionadoMov1 
         Height          =   240
         Left            =   135
         TabIndex        =   34
         Top             =   525
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid GridMov1 
         Height          =   3495
         Left            =   30
         TabIndex        =   40
         Top             =   795
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   16
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin MSMask.MaskEdBox NumRefExterna1 
         Height          =   240
         Left            =   6870
         TabIndex        =   39
         Top             =   555
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataMov1 
         Height          =   240
         Left            =   930
         TabIndex        =   36
         Top             =   540
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ValorMov1 
         Height          =   240
         Left            =   1965
         TabIndex        =   37
         Top             =   540
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Total2 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7590
         TabIndex        =   65
         Top             =   150
         Width           =   1485
      End
      Begin VB.Label Label10 
         Caption         =   "Total:"
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
         Left            =   6960
         TabIndex        =   64
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label8 
         Caption         =   "Movimentos no Sistema:"
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
         TabIndex        =   59
         Top             =   555
         Width           =   4260
      End
      Begin VB.Label Label9 
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
         Height          =   240
         Left            =   375
         TabIndex        =   60
         Top             =   180
         Width           =   1410
      End
   End
   Begin VB.CommandButton BotaoFechar 
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
      Left            =   8175
      Picture         =   "ConciliacaoBancariaOcx.ctx":3699
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   45
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5790
      Left            =   75
      TabIndex        =   53
      Top             =   120
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   10213
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Extrato CNAB"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Extrato em Papel"
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
Attribute VB_Name = "ConciliacaoBancariaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'??? transferir p/global
Private Const STRING_TIPOSMOVTOCTACORRENTE_NOMERED = 50

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGrid1 As AdmGrid 'extrato
Dim objGrid2 As AdmGrid 'movCCI do frame extrato papel GridMov1
Dim objGrid3 As AdmGrid 'movCCI do frame extrato CNAB GridMov
Dim iFrameAtual As Integer
Dim iFrameTrabalho As Integer '0=não há frame selecionado 2=FRAME_CNAB 3=FRAME_PAPEL

'Última expressão de seleção utilizada nos frames CNAB e Papel
Dim sSQL_CNAB As String
Dim sSQL_Papel As String

'Frames da tela
Const FRAME_SELECAO = 1
Const FRAME_CNAB = 2
Const FRAME_PAPEL = 3

'Tipos de ordenação dos grids
Const ORDENACAO_DATA = 1
Const ORDENACAO_VALOR = 2
Const ORDENACAO_HISTORICO = 3

'Tipos de pesquisa dos grids
Const PROCURA_DATA = 1
Const PROCURA_VALOR = 2
Const PROCURA_HISTORICO = 3

'Colunas dos grids
Const GRID_SELECIONADO_COL = 1
Const GRID_CONCILIADO_COL = 2

Dim iGridDataCol As Integer
Dim iGridValorCol As Integer
Dim iGridHistoricoCol As Integer
Dim iGridNumRefExternaCol As Integer
Dim iGridCategoriaCol As Integer
Dim iGridCodLctoBancoCol As Integer
Dim iGridDocumentoCol As Integer

'Indica se a seleção está pedindo para exibir todos os registros ou somente os não conciliados
Dim iExibeTodosExt As Integer
Dim iExibeTodosMov1 As Integer

'Indicação de conciliação no grid
Const S_CONCILIADO As String = "1"
Const S_NAO_CONCILIADO As String = "0"

'Indicação de marcação no grid
Const S_MARCADO As String = "1"
Const S_DESMARCADO As String = "0"

'Colecões que armazenam os movimentos de conta corrente e lançamentos oriundos dos extratos
Dim colMovCCI As New Collection
Dim colMovCCI1 As New Collection
Dim colExtrato As New Collection

Dim gcolTiposMovCta As New AdmColCodigoNome

Private Sub Botao_Conciliar_Click()

Dim colGridExtMarcado As New Collection
Dim colGridMovMarcado As New Collection
Dim iLinha As Integer
Dim vIndice As Variant
Dim iIndiceExt As Integer
Dim iIndiceMov As Integer
Dim dTotalMov As Double
Dim dTotalExt As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_Botao_Conciliar_Click

    dTotalExt = 0

    'cria uma coleção com as linhas do grid extrato marcadas
    For iLinha = 1 To objGrid1.iLinhasExistentes
        If GridExtrato.TextMatrix(iLinha, GRID_SELECIONADO_COL) = S_MARCADO Then
            colGridExtMarcado.Add iLinha
            dTotalExt = dTotalExt + CDbl(GridExtrato.TextMatrix(iLinha, iGridValorCol))
        End If
    Next
    
    dTotalMov = 0
    
    'cria uma coleção com as linhas do grid de movimentos marcadas
    For iLinha = 1 To objGrid3.iLinhasExistentes
        If GridMov.TextMatrix(iLinha, GRID_SELECIONADO_COL) = S_MARCADO Then
            colGridMovMarcado.Add iLinha
            dTotalMov = dTotalMov + CDbl(GridMov.TextMatrix(iLinha, iGridValorCol))
        End If
    Next

    'os dois grids não podem ter mais de um elemento selecionado ao mesmo tempo
    If colGridExtMarcado.Count > 1 And colGridMovMarcado.Count > 1 Then Error 10809

    'verifica se foi marcado algum lançamento de extrato
    If colGridExtMarcado.Count = 0 Then Error 10810

    'verifica se foi marcado algum movimento de conta corrente
    If colGridMovMarcado.Count = 0 Then Error 10811

    'verifica se o total dos movimentos selecionados coincide com o total dos extratos
    If dTotalMov <> dTotalExt Then
    
        'se os totais não coincidirem, pede confirmação da operação
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONCILIACAO_TOTAL_MOV_EXT_DIFERENTES", dTotalMov, dTotalExt)
        
        If vbMsgRes = vbNo Then Exit Sub
        
    End If

    'processa a conciliação
    lErro = CF("ConciliacaoBancaria_Grava", colExtrato, colGridExtMarcado, colMovCCI, colGridMovMarcado, iIndiceExt, iIndiceMov)
    If lErro <> SUCESSO And lErro <> 10818 And lErro <> 10821 Then Error 10826

    If lErro = 10818 Then Error 10827
    
    If lErro = 10821 Then Error 10828
    
    'se somente são exibidos os não conciliados ==> remove os conciliados da tela
    If Todos.Value = False Then
            
        For Each vIndice In colGridExtMarcado
            colExtrato.Remove vIndice
        Next
        
        For Each vIndice In colGridMovMarcado
            colMovCCI.Remove vIndice
        Next
            
        lErro = Preenche_GridMov(colMovCCI)
        If lErro <> SUCESSO Then Error 10873
    
        lErro = Preenche_GridExtrato(colExtrato)
        If lErro <> SUCESSO Then Error 10874
            
    Else
    
        For Each vIndice In colGridMovMarcado
            GridMov.TextMatrix(vIndice, GRID_SELECIONADO_COL) = S_DESMARCADO
            GridMov.TextMatrix(vIndice, GRID_CONCILIADO_COL) = S_CONCILIADO
        Next
    
        Call Grid_Refresh_Checkbox(objGrid3)
    
        For Each vIndice In colGridExtMarcado
            GridExtrato.TextMatrix(vIndice, GRID_SELECIONADO_COL) = S_DESMARCADO
            GridExtrato.TextMatrix(vIndice, GRID_CONCILIADO_COL) = S_CONCILIADO
        Next
            
        Call Grid_Refresh_Checkbox(objGrid1)
            
    End If
    
    Exit Sub
    
Erro_Botao_Conciliar_Click:

    Select Case Err
    
        Case 10809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDS_MAIS_UM_ELEMENTO_SELECIONADO", Err)
    
        Case 10810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_EXTRATO_SEM_SELECAO", Err)
    
        Case 10811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_MOV_SEM_SELECAO", Err)
    
        Case 10826, 10873, 10874
    
        Case 10827
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_EXT_JA_CONCILIADO", Err, iIndiceExt)
        
        Case 10828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_MOV_JA_CONCILIADO", Err, iIndiceMov)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154507)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Botao_ConciliarMov1_Click()

Dim colGridMov1Marcado As New Collection
Dim iLinha As Integer
Dim vIndice As Variant
Dim iIndiceMov As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_Botao_ConciliarMov1_Click

    'cria uma coleção com as linhas do grid de movimentos marcadas
    For iLinha = 1 To objGrid2.iLinhasExistentes
        If GridMov1.TextMatrix(iLinha, GRID_SELECIONADO_COL) = S_MARCADO Then
            colGridMov1Marcado.Add iLinha
        End If
    Next

    'verifica se foi marcado algum movimento de conta corrente
    If colGridMov1Marcado.Count = 0 Then Error 10882

    'processa a conciliação
    lErro = CF("MovCCI_Atualiza_Conciliado", colMovCCI1, colGridMov1Marcado, iIndiceMov)
    If lErro <> SUCESSO And lErro <> 10887 Then Error 10892

    If lErro = 10887 Then Error 10893
    
    'se somente são exibidos os não conciliados ==> remove os conciliados da tela
    If Todos.Value = False Then
            
        For iLinha = colGridMov1Marcado.Count To 1 Step -1
            colMovCCI1.Remove colGridMov1Marcado.Item(iLinha)
        Next
            
        lErro = Preenche_GridMov1(colMovCCI1)
        If lErro <> SUCESSO Then Error 10894
    
    Else
    
        For Each vIndice In colGridMov1Marcado
            GridMov1.TextMatrix(vIndice, GRID_SELECIONADO_COL) = S_DESMARCADO
            GridMov1.TextMatrix(vIndice, GRID_CONCILIADO_COL) = S_CONCILIADO
        Next
        
        Call Grid_Refresh_Checkbox(objGrid2)
        
    End If
    
    Exit Sub
    
Erro_Botao_ConciliarMov1_Click:

    Select Case Err
    
        Case 10882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_MOV_SEM_SELECAO", Err)
    
        Case 10893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_MOV_JA_CONCILIADO", Err, iIndiceMov)
            
        Case 10894
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154508)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Botao_Desconciliar_Click()

Dim colGridExtMarcado As New Collection
Dim colGridMovMarcado As New Collection
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Botao_Desconciliar_Click

    'cria uma coleção com as linhas do grid extrato marcadas
    For iLinha = 1 To objGrid1.iLinhasExistentes
        If GridExtrato.TextMatrix(iLinha, GRID_SELECIONADO_COL) = S_MARCADO Then
            colGridExtMarcado.Add iLinha
        End If
    Next
    
    'cria uma coleção com as linhas do grid de movimentos marcadas
    For iLinha = 1 To objGrid3.iLinhasExistentes
        If GridMov.TextMatrix(iLinha, GRID_SELECIONADO_COL) = S_MARCADO Then
            colGridMovMarcado.Add iLinha
        End If
    Next

    'os dois grids não podem ter nenhum elemento selecionado ao mesmo tempo
    If colGridExtMarcado.Count = 0 And colGridMovMarcado.Count = 0 Then Error 10829

    lErro = CF("ConciliacaoBancaria_Exclui", colExtrato, colGridExtMarcado, colMovCCI, colGridMovMarcado)
    If lErro <> SUCESSO Then Error 10830

    Call Ordenados_Click

    Exit Sub
    
Erro_Botao_Desconciliar_Click:

    Select Case Err
    
        Case 10829
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_EXTRATO_MOV_SEM_SELECAO", Err)
    
        Case 10830
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154509)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Botao_DesconciliarMov1_Click()

Dim colGridExtMarcado As New Collection
Dim colGridMovMarcado As New Collection
Dim iLinha As Integer
Dim vIndice As Variant
Dim lErro As Long

On Error GoTo Erro_Botao_DesconciliarMov1_Click

    'cria uma coleção com as linhas do grid de movimentos marcadas
    For iLinha = 1 To objGrid2.iLinhasExistentes
        If GridMov1.TextMatrix(iLinha, GRID_SELECIONADO_COL) = S_MARCADO Then
            colGridMovMarcado.Add iLinha
        End If
    Next

    'o grid não pode ter nenhum elemento selecionado
    If colGridMovMarcado.Count = 0 Then Error 10910

    lErro = CF("ConciliacaoBancaria_Exclui", colExtrato, colGridExtMarcado, colMovCCI1, colGridMovMarcado)
    If lErro <> SUCESSO Then Error 10911

    For Each vIndice In colGridMovMarcado
        GridMov1.TextMatrix(vIndice, GRID_SELECIONADO_COL) = S_DESMARCADO
        GridMov1.TextMatrix(vIndice, GRID_CONCILIADO_COL) = S_NAO_CONCILIADO
    Next
        
    Call Grid_Refresh_Checkbox(objGrid2)

    Exit Sub
    
Erro_Botao_DesconciliarMov1_Click:

    Select Case Err
    
        Case 10910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_MOV_SEM_SELECAO", Err)
    
        Case 10911
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154510)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Botao_ExibirMov_Click()

Dim iLinha As Integer
Dim objExtrBcoDet As ClassExtrBcoDet
Dim lErro As Long
Dim colConciliacao As New Collection
Dim objConciliacaoBancaria As ClassConciliacaoBancaria
Dim objMovCCI As ClassMovContaCorrente
Dim iIndice As Integer

On Error GoTo Erro_Botao_ExibirMov_Click

    'verifica se alguma linha está selecionado
    If GridExtrato.Row < 1 Then Error 10864

    Set objExtrBcoDet = colExtrato.Item(GridExtrato.Row)
    
    lErro = CF("ConciliacaoBancaria_Le_Ext", objExtrBcoDet.iCodConta, objExtrBcoDet.iNumExtrato, objExtrBcoDet.lSeqLcto, colConciliacao)
    If lErro <> SUCESSO Then Error 10865
    
    If colConciliacao.Count = 0 Then Error 10872
    
    'Limpa as marcações existentes no grid
    For iLinha = 1 To objGrid3.iLinhasExistentes
        GridMov.TextMatrix(iLinha, GRID_SELECIONADO_COL) = S_DESMARCADO
    Next
    
    iLinha = 0
    
    For Each objConciliacaoBancaria In colConciliacao
    
        For iIndice = 1 To colMovCCI.Count
        
            Set objMovCCI = colMovCCI.Item(iIndice)
        
            If objConciliacaoBancaria.iCodConta = objMovCCI.iCodConta And objConciliacaoBancaria.lSequencialMovto = objMovCCI.lSequencial Then
            
                GridMov.TextMatrix(iIndice, GRID_SELECIONADO_COL) = S_MARCADO
                
                If iLinha = 0 Then iLinha = iIndice
                
                Exit For
                
            End If
    
        Next
    Next
    
    GridExtrato.TextMatrix(GridExtrato.Row, GRID_SELECIONADO_COL) = S_MARCADO
    
    
    Call Grid_Refresh_Checkbox(objGrid3)
    Call Grid_Refresh_Checkbox(objGrid1)

    GridMov.TopRow = iLinha

    Exit Sub
    
Erro_Botao_ExibirMov_Click:

    Select Case Err
    
        Case 10864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_EXT_SEM_SELECAO", Err)
            
        Case 10865
            
        Case 10872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXT_SEM_MOV_CONCILIADO", Err, objExtrBcoDet.iCodConta, objExtrBcoDet.iNumExtrato, objExtrBcoDet.lSeqLcto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154511)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Botao_ExibirExt_Click()

Dim iLinha As Integer
Dim objExtrBcoDet As ClassExtrBcoDet
Dim lErro As Long
Dim colConciliacao As New Collection
Dim objConciliacaoBancaria As ClassConciliacaoBancaria
Dim objMovCCI As ClassMovContaCorrente
Dim iIndice As Integer

On Error GoTo Erro_Botao_ExibirExt_Click

    'verifica se alguma linha está selecionada
    If GridMov.Row < 1 Then Error 10875

    Set objMovCCI = colMovCCI.Item(GridMov.Row)
    
    lErro = CF("ConciliacaoBancaria_Le_Mov", objMovCCI.iCodConta, objMovCCI.lSequencial, colConciliacao)
    If lErro <> SUCESSO Then Error 10876
    
    If colConciliacao.Count = 0 Then Error 10877
    
    'Limpa as marcações existentes no grid
    For iLinha = 1 To objGrid1.iLinhasExistentes
        GridExtrato.TextMatrix(iLinha, GRID_SELECIONADO_COL) = S_DESMARCADO
    Next
    
    iLinha = 0
    
    For Each objConciliacaoBancaria In colConciliacao
    
        For iIndice = 1 To colExtrato.Count
        
            Set objExtrBcoDet = colExtrato.Item(iIndice)
        
            If objConciliacaoBancaria.iCodConta = objExtrBcoDet.iCodConta And objConciliacaoBancaria.iNumExtrato = objExtrBcoDet.iNumExtrato And objConciliacaoBancaria.lSeqExtrBco = objExtrBcoDet.lSeqLcto Then
            
                GridExtrato.TextMatrix(iIndice, GRID_SELECIONADO_COL) = S_MARCADO
                
                If iLinha = 0 Then iLinha = iIndice
                
                Exit For
                
            End If
    
        Next
    Next
    
    GridMov.TextMatrix(GridMov.Row, GRID_SELECIONADO_COL) = S_MARCADO
    
    Call Grid_Refresh_Checkbox(objGrid1)
    Call Grid_Refresh_Checkbox(objGrid3)
    
    GridExtrato.TopRow = iLinha

    Exit Sub
    
Erro_Botao_ExibirExt_Click:

    Select Case Err
    
        Case 10875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_MOV_SEM_SELECAO", Err)
            
        Case 10876
            
        Case 10877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOV_SEM_EXT_CONCILIADO", Err, objMovCCI.iCodConta, objMovCCI.lSequencial)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154512)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Botao_ProcurarMov_Click()
    
Dim objExtrBcoDet As ClassExtrBcoDet
Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objMovCCI As ClassMovContaCorrente
Dim iAchou As Integer

On Error GoTo Erro_Botao_ProcurarMov_Click

    'verifica se alguma linha está selecionada
    If GridExtrato.Row < 1 Then Error 10895
    
    Set objExtrBcoDet = colExtrato.Item(GridExtrato.Row)
    
    iLinha = GridMov.Row
    
    iAchou = 0
    
    Select Case Procura.ItemData(Procura.ListIndex)
    
        Case PROCURA_DATA
                                    
            For iIndice = 1 To colMovCCI.Count
                                
                If iLinha = colMovCCI.Count Then iLinha = 0
                
                iLinha = iLinha + 1
                
                Set objMovCCI = colMovCCI.Item(iLinha)
                
                If objExtrBcoDet.dtData = objMovCCI.dtDataMovimento Then
                    GridMov.Row = iLinha
                    iAchou = 1
                    Exit For
                End If
            Next
            
            If iAchou = 0 Then Error 10896
            
            
        Case PROCURA_VALOR
                                    
            For iIndice = 1 To colMovCCI.Count
                                
                If iLinha = colMovCCI.Count Then iLinha = 0
                
                iLinha = iLinha + 1
                
                Set objMovCCI = colMovCCI.Item(iLinha)
                
                If Abs(objExtrBcoDet.dValor - objMovCCI.dValor) < DELTA_VALORMONETARIO Then
                    GridMov.Row = iLinha
                    iAchou = 1
                    Exit For
                End If
            Next
            
            If iAchou = 0 Then Error 10897
            
        Case PROCURA_HISTORICO
                                    
            For iIndice = 1 To colMovCCI.Count
                                
                If iLinha = colMovCCI.Count Then iLinha = 0
                
                iLinha = iLinha + 1
                
                Set objMovCCI = colMovCCI.Item(iLinha)
                
                If objExtrBcoDet.sHistorico = objMovCCI.sHistorico Then
                    GridMov.Row = iLinha
                    iAchou = 1
                    Exit For
                End If
            Next
            
            If iAchou = 0 Then Error 10898
                    
    End Select
    
    
    If GridExtrato.Col <> iGridDataCol - 1 Or GridExtrato.RowSel <> GridExtrato.Row Or GridExtrato.ColSel <> GridExtrato.Cols - 1 Then
        GridExtrato.Col = iGridDataCol - 1
        GridExtrato.RowSel = GridExtrato.Row
        GridExtrato.ColSel = GridExtrato.Cols - 1
    End If
    
    GridMov.TopRow = GridMov.Row
    GridMov.Col = iGridDataCol - 1
    GridMov.RowSel = GridMov.Row
    GridMov.ColSel = GridMov.Cols - 1
    
    Exit Sub
    
Erro_Botao_ProcurarMov_Click:

    Select Case Err
    
        Case 10895
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_EXT_SEM_SELECAO", Err)
            
        Case 10896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_GRID_MOV_DATA", Err, CStr(objExtrBcoDet.dtData))
            
        Case 10897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_GRID_MOV_VALOR", Err, objExtrBcoDet.dValor)
            
        Case 10898
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_GRID_MOV_HISTORICO", Err, objExtrBcoDet.sHistorico)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154513)
            
    End Select
    
    Exit Sub
            
End Sub

Private Sub Botao_ProcurarExt_Click()
    
Dim objExtrBcoDet As ClassExtrBcoDet
Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objMovCCI As ClassMovContaCorrente
Dim iAchou As Integer

On Error GoTo Erro_Botao_ProcurarExt_Click

    'verifica se alguma linha está selecionada
    If GridMov.Row < 1 Then Error 10899
    
    Set objMovCCI = colMovCCI.Item(GridMov.Row)
    
    iLinha = GridExtrato.Row
    
    iAchou = 0
    
    Select Case Procura.ItemData(Procura.ListIndex)
    
        Case PROCURA_DATA
                                    
            For iIndice = 1 To colExtrato.Count
                                
                If iLinha = colExtrato.Count Then iLinha = 0
                
                iLinha = iLinha + 1
                
                Set objExtrBcoDet = colExtrato.Item(iLinha)
                
                If objExtrBcoDet.dtData = objMovCCI.dtDataMovimento Then
                    GridExtrato.Row = iLinha
                    iAchou = 1
                    Exit For
                End If
            Next
            
            If iAchou = 0 Then Error 10900
            
        Case PROCURA_VALOR
                                    
            For iIndice = 1 To colExtrato.Count
                                
                If iLinha = colExtrato.Count Then iLinha = 0
                
                iLinha = iLinha + 1
                
                Set objExtrBcoDet = colExtrato.Item(iLinha)
                
                If objExtrBcoDet.dValor = objMovCCI.dValor Then
                    GridExtrato.Row = iLinha
                    iAchou = 1
                    Exit For
                End If
            Next
            
            If iAchou = 0 Then Error 10901
            
        Case PROCURA_HISTORICO
                                    
            For iIndice = 1 To colExtrato.Count
                                
                If iLinha = colExtrato.Count Then iLinha = 0
                
                iLinha = iLinha + 1
                
                Set objExtrBcoDet = colExtrato.Item(iLinha)
                
                If objExtrBcoDet.sHistorico = objMovCCI.sHistorico Then
                    GridExtrato.Row = iLinha
                    iAchou = 1
                    Exit For
                End If
            Next
            
            If iAchou = 0 Then Error 10902
                    
    End Select
    
    If GridMov.Col <> iGridDataCol - 1 Or GridMov.RowSel <> GridMov.Row Or GridMov.ColSel <> GridMov.Cols - 1 Then
        GridMov.Col = iGridDataCol - 1
        GridMov.RowSel = GridMov.Row
        GridMov.ColSel = GridMov.Cols - 1
    End If
    
    GridExtrato.TopRow = GridExtrato.Row
    GridExtrato.Col = iGridDataCol - 1
    GridExtrato.RowSel = GridExtrato.Row
    GridExtrato.ColSel = GridExtrato.Cols - 1
        
    Exit Sub
    
Erro_Botao_ProcurarExt_Click:

    Select Case Err
    
        Case 10899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_MOV_SEM_SELECAO", Err)
            
        Case 10900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_GRID_EXT_DATA", Err, CStr(objMovCCI.dtDataMovimento))
            
        Case 10901
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_GRID_EXT_VALOR", Err, objMovCCI.dValor)
            
        Case 10902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_GRID_EXT_HISTORICO", Err, objMovCCI.sHistorico)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154514)
            
    End Select
    
    Exit Sub
            
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    sSQL_CNAB = ""
    sSQL_Papel = ""
    
    iFrameAtual = FRAME_SELECAO
    iFrameTrabalho = 0
    
    'Carrega a combo com o codigo e nome das contas correntes
    lErro = Carrega_CodContaCorrente()
    If lErro <> SUCESSO Then gError 10734
    
    lErro = Carrega_TiposMovCta()
    If lErro <> SUCESSO Then gError 96059
    
    DataDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'inicializa as comboboxes de ordenação dos grids
    For iIndice = 0 To Ordenados.ListCount - 1
        If Ordenados.ItemData(iIndice) = ORDENACAO_DATA Then
            Ordenados.ListIndex = iIndice
            Ordenados1.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'inicializa a combobox de pesquisa dos grids
    For iIndice = 0 To Procura.ListCount - 1
        If Procura.ItemData(iIndice) = PROCURA_DATA Then
            Procura.ListIndex = iIndice
            Exit For
        End If
    Next
    
    lErro = Inicializa_GridExtrato()
    If lErro <> SUCESSO Then gError 10803
    
    lErro = Inicializa_GridMov()
    If lErro <> SUCESSO Then gError 10804
    
    lErro = Inicializa_GridMov1()
    If lErro <> SUCESSO Then gError 10805
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 10734, 10803, 10804, 10805, 96059
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154515)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing
    Set objGrid2 = Nothing
    Set objGrid3 = Nothing

    Set colMovCCI = Nothing
    Set colMovCCI1 = Nothing
    Set colExtrato = Nothing

    Set gcolTiposMovCta = Nothing

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub
        
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
    
        If Opcao.SelectedItem.Index = FRAME_CNAB Then
        
            Parent.HelpContextID = IDH_CONCILIACAO_EXTRATO_CNAB
            
            'Exibe os dados nos grids
            lErro = Processa_Selecao_CNAB()
            If lErro <> SUCESSO Then Error 10744
            
            iFrameTrabalho = FRAME_CNAB
            
        ElseIf Opcao.SelectedItem.Index = FRAME_PAPEL Then
        
            Parent.HelpContextID = IDH_CONCILIACAO_BANCARIA_EXTRATO_PAPEL
            
            'Exibe os movimentos de conta corrente no grid
            lErro = Processa_Selecao_Papel()
            If lErro <> SUCESSO Then Error 10775
            
            iFrameTrabalho = FRAME_PAPEL
            
        ElseIf Opcao.SelectedItem.Index = FRAME_SELECAO Then
            
            Parent.HelpContextID = IDH_CONCILIACAO_BANCARIA_SELECAO
        
        End If
    
    End If

    Exit Sub
    
Erro_Opcao_Click:

    Select Case Err

        Case 10744, 10775
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154516)
        
    End Select
    
    Exit Sub

End Sub

Private Function Processa_Selecao_CNAB() As Long
'exibe os dados de movimento de conta corrente e os lançamentos de extrato nos respectivos grids

Dim sSelecaoMov As String
Dim sSelecaoExt As String
Dim sOrdenacaoMov As String
Dim sOrdenacaoExt As String
Dim avCampo(1 To 6) As Variant
Dim iNumCampo As Integer
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Processa_Selecao_CNAB

    lErro = Monta_Selecao(sSelecaoMov, sSelecaoExt, avCampo(), iNumCampo)
    If lErro <> SUCESSO Then gError 10762

    'se a selecao atual for diferente da que está armazenada
'    If sSelecaoMov <> sSQL_CNAB Or iFrameTrabalho <> FRAME_CNAB Then

        lErro = Monta_Ordenacao(sOrdenacaoMov, sOrdenacaoExt, Ordenados)
        If lErro <> SUCESSO Then gError 10763
    
        Set colMovCCI = New Collection
    
        lErro = CF("MovContaCorrente_Le_Conciliacao", sSelecaoMov, avCampo(), iNumCampo, colMovCCI, sOrdenacaoMov)
        If lErro <> SUCESSO And lErro <> 10757 And lErro <> 10758 Then gError 10761
    
        'Se não existirem lançamentos de extrato para a seleção em questão
        If lErro = 10757 Then gError 10764
        
        If lErro = 10758 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_MOV_ULTRAPASSOU_LIMITE")
    
        lErro = Gera_HistoricoMovto(colMovCCI)
        If lErro <> SUCESSO Then gError 96061
        
        Set colExtrato = New Collection
    
        lErro = CF("LctosExtratoBancario_Le_Conciliacao", sSelecaoExt, avCampo(), iNumCampo, colExtrato, sOrdenacaoExt)
        If lErro <> SUCESSO And lErro <> 10768 And lErro <> 10769 Then gError 10780
    
        'se não existirem lançamentos de extrato para a seleção em questão
        If lErro = 10768 Then gError 10772
        
        If lErro = 10769 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_LANCEXTRATO_ULTRAPASSOU_LIMITE")
    
        If iExibeTodosExt <> Todos.Value Then
            
            lErro = Inicializa_GridExtrato()
            If lErro <> SUCESSO Then gError 10806
    
            lErro = Inicializa_GridMov()
            If lErro <> SUCESSO Then gError 10807
    
        End If
    
        lErro = Preenche_GridMov(colMovCCI)
        If lErro <> SUCESSO Then gError 10773
    
        lErro = Preenche_GridExtrato(colExtrato)
        If lErro <> SUCESSO Then gError 10774
    
        Botao_Desconciliar.Enabled = Todos.Value
    
        sSQL_CNAB = sSelecaoMov
        
'    End If
            
    Processa_Selecao_CNAB = SUCESSO

    Exit Function

Erro_Processa_Selecao_CNAB:

    Processa_Selecao_CNAB = gErr

    Select Case gErr

        Case 10761, 10762, 10763, 10773, 10774, 10780, 10806, 10807
    
        Case 10764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOS_INEXISTENTES_CONCILIACAO", gErr)
    
        Case 10772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCEXTRATO_INEXISTENTES_CONCILIACAO", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154517)

    End Select

    Exit Function

End Function

Private Function Processa_Selecao_Papel() As Long
'Exibe os dados de movimento de conta corrente no grid

Dim lErro As Long
Dim sSelecaoMov As String
Dim sSelecaoExt As String
Dim sOrdenacaoMov As String
Dim sOrdenacaoExt As String
Dim avCampo(1 To 6) As Variant
Dim iNumCampo As Integer

On Error GoTo Erro_Processa_Selecao_Papel

    lErro = Monta_Selecao(sSelecaoMov, sSelecaoExt, avCampo(), iNumCampo)
    If lErro <> SUCESSO Then gError 10776

    'se a selecao atual for diferente da que está armazenada
'    If sSelecaoMov <> sSQL_Papel Or iFrameTrabalho <> FRAME_PAPEL Then

        lErro = Monta_Ordenacao(sOrdenacaoMov, sOrdenacaoExt, Ordenados1)
        If lErro <> SUCESSO Then gError 10777
    
        Set colMovCCI1 = New Collection
    
        lErro = CF("MovContaCorrente_Le_Conciliacao", sSelecaoMov, avCampo(), iNumCampo, colMovCCI1, sOrdenacaoMov)
        If lErro <> SUCESSO And lErro <> 10757 And lErro <> 10758 Then gError 10778
    
        'se não existirem lançamentos de extrato para a seleção em questão
        If lErro = 10757 Then
            'Maristela - Acrescentei o Call Grid_Limpa(objGrid2)
            Call Grid_Limpa(objGrid2)
            Error 10779
        End If
                
        If lErro = 10758 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_MOV_ULTRAPASSOU_LIMITE")
    
        lErro = Gera_HistoricoMovto(colMovCCI1)
        If lErro <> SUCESSO Then gError 96061

        If iExibeTodosMov1 <> Todos.Value Then
            
            lErro = Inicializa_GridMov1()
            If lErro <> SUCESSO Then gError 10808
    
        End If
    
        lErro = Preenche_GridMov1(colMovCCI1)
        If lErro <> SUCESSO Then gError 10781
    
        Botao_DesconciliarMov1.Enabled = Todos.Value
    
        sSQL_Papel = sSelecaoMov
        
'    End If
            
    Processa_Selecao_Papel = SUCESSO

    Exit Function

Erro_Processa_Selecao_Papel:

    Processa_Selecao_Papel = gErr

    Select Case gErr

        Case 10776, 10777, 10778, 10781, 10808
    
        Case 10779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOS_INEXISTENTES_CONCILIACAO", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154518)

    End Select

    Exit Function

End Function

Function Gera_HistoricoMovto(colMovCCI As Collection) As Long
'gera historico "automatico"

Dim lErro As Long
Dim sHistoricoAuto As String
Dim objCodigoNome As AdmCodigoNome
Dim objMovCCI As ClassMovContaCorrente

On Error GoTo Erro_Gera_HistoricoMovto

    For Each objMovCCI In colMovCCI
    
        If objMovCCI.sHistorico = "" Then
    
            For Each objCodigoNome In gcolTiposMovCta
                If objCodigoNome.iCodigo = objMovCCI.iTipo Then
                    sHistoricoAuto = objCodigoNome.sNome
                    Exit For
                End If
            Next
   
        Else
        
            sHistoricoAuto = objMovCCI.sHistorico
        
        End If
                
'            If (objMovCCI.iTipoMeioPagto = Cheque Or objMovCCI.iTipoMeioPagto = BORDERO) And objMovCCI.lNumero <> 0 Then
        Select Case objMovCCI.iTipoMeioPagto
        
            Case Cheque
            
                If InStr(UCase(sHistoricoAuto), "CHEQUE") = 0 Then sHistoricoAuto = sHistoricoAuto & " Cheque "
                
            Case BORDERO
            
                If InStr(UCase(sHistoricoAuto), "BORDER") = 0 Then sHistoricoAuto = sHistoricoAuto & " Borderô "
        
        End Select
        
        If objMovCCI.lNumero <> 0 Then
            objMovCCI.sHistorico = sHistoricoAuto & " No. " & objMovCCI.lNumero
        Else
            objMovCCI.sHistorico = sHistoricoAuto
        End If
            
    Next
    
    Gera_HistoricoMovto = SUCESSO
        
Exit Function

Erro_Gera_HistoricoMovto:
   
    Gera_HistoricoMovto = gErr
   
    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154519)

    End Select
    
    Exit Function
    
End Function

Function Monta_Selecao(sSelecaoMov As String, sSelecaoExt As String, avCampo() As Variant, iNumCampo As Integer) As Long
'monta a expressão de seleção SQL

Dim lErro As Long

On Error GoTo Erro_Monta_Selecao

    'Verifica se a conta foi preenchida
    If Len(Trim(CodCCI.Text)) = 0 Then Error 10745

    sSelecaoMov = "CodConta = ?"
    sSelecaoExt = "CodConta = ?"
    
    avCampo(1) = Codigo_Extrai(CodCCI.Text)
    iNumCampo = 1
    
    If Len(Trim(DataDe.ClipText)) > 0 Then
        
        sSelecaoMov = sSelecaoMov + " AND DataMovimento >= ?"
        sSelecaoExt = sSelecaoExt + " AND Data >= ?"
    
        iNumCampo = iNumCampo + 1
        avCampo(iNumCampo) = CDate(DataDe.Text)
        
    End If
    
    If Len(Trim(DataAte.ClipText)) > 0 Then
        
        sSelecaoMov = sSelecaoMov + " AND DataMovimento <= ?"
        sSelecaoExt = sSelecaoExt + " AND Data <= ?"
    
        iNumCampo = iNumCampo + 1
        avCampo(iNumCampo) = CDate(DataAte.Text)
        
        If Len(Trim(DataDe.ClipText)) > 0 Then
            If CDate(DataAte.Text) < CDate(DataDe.Text) Then Error 10746
        End If
        
    End If
    
        
    If Len(Trim(ValorDe.Text)) > 0 Then
        
        sSelecaoMov = sSelecaoMov + " AND Valor >= ?"
        sSelecaoExt = sSelecaoExt + " AND Valor >= ?"
    
        iNumCampo = iNumCampo + 1
        avCampo(iNumCampo) = CDbl(ValorDe.Text)
        
    End If
        
    If Len(Trim(ValorAte.Text)) > 0 Then
        
        sSelecaoMov = sSelecaoMov + " AND Valor <= ?"
        sSelecaoExt = sSelecaoExt + " AND Valor <= ?"
    
        iNumCampo = iNumCampo + 1
        avCampo(iNumCampo) = CDbl(ValorAte.Text)
        
        If Len(Trim(ValorDe.ClipText)) > 0 Then
            If CDbl(ValorAte.Text) < CDbl(ValorDe.Text) Then Error 10747
        End If
        
    End If
        
    If NaoConciliados.Value = True Then
        
        sSelecaoMov = sSelecaoMov + " AND Conciliado = ?"
        sSelecaoExt = sSelecaoExt + " AND Conciliado = ?"
    
        iNumCampo = iNumCampo + 1
        avCampo(iNumCampo) = CInt(NAO_CONCILIADO)
        
    End If

    Monta_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Selecao:

    Monta_Selecao = Err

    Select Case Err
    
        Case 10745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 10746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case 10747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154520)

    End Select

    Exit Function

End Function

Private Function Monta_Ordenacao(sOrdenacaoMov As String, sOrdenacaoExt As String, Ordenacao As ComboBox) As Long
'monta a expressão de ordenação SQL

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Ordenacao

    Select Case Ordenacao.ItemData(Ordenacao.ListIndex)
    
        Case ORDENACAO_DATA
        
            sOrdenacaoMov = " ORDER BY DataMovimento"
            sOrdenacaoExt = " ORDER BY Data"
            
        Case ORDENACAO_VALOR
        
            sOrdenacaoMov = " ORDER BY Valor"
            sOrdenacaoExt = " ORDER BY Valor"
        
        Case ORDENACAO_HISTORICO
        
            sOrdenacaoMov = " ORDER BY Historico"
            sOrdenacaoExt = " ORDER BY Historico"

    End Select

    Monta_Ordenacao = SUCESSO
    
    Exit Function

Erro_Monta_Ordenacao:

    Monta_Ordenacao = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154521)

    End Select

    Exit Function

End Function

Private Function Carrega_CodContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_CodContaCorrente

    'Leitura dos códigos e descrições das Contas existentes no BD
    lErro = CF("ContasCorrentes_Bancarias_Le_CodigosNomesRed", colCodigoDescricao)
    If lErro <> SUCESSO Then Error 10735

    For Each objCodigoNome In colCodigoDescricao

        'Insere na combo de contas correntes
        CodCCI.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodCCI.ItemData(CodCCI.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 10735

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154522)

    End Select

    Exit Function

End Function

Private Function Carrega_TiposMovCta() As Long
'Carrega os tipos de movimento

Dim lErro As Long

On Error GoTo Erro_Carrega_TiposMovCta

    'Leitura dos códigos e NomeReduzido dos tipos de movimentos existentes no BD
    lErro = CF("Cod_Nomes_Le", "tiposmovtoctacorrente", "Codigo", "NomeReduzido", STRING_TIPOSMOVTOCTACORRENTE_NOMERED, gcolTiposMovCta)
    If lErro <> SUCESSO Then gError 96060

    Carrega_TiposMovCta = SUCESSO
    
    Exit Function

Erro_Carrega_TiposMovCta:

    Carrega_TiposMovCta = gErr

    Select Case gErr

        Case 96060

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154523)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridExtrato() As Long
   
    iExibeTodosExt = Todos.Value
    
    Set objGrid1 = New AdmGrid
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    objGrid1.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid1.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'titulos do grid
    objGrid1.colColuna.Add ("")
    objGrid1.colColuna.Add ("Marcado")
    If Todos.Value = True Then objGrid1.colColuna.Add ("Conciliado")
    objGrid1.colColuna.Add ("Data")
    objGrid1.colColuna.Add ("Valor")
    objGrid1.colColuna.Add ("Histórico")
    objGrid1.colColuna.Add ("Cat.")
    objGrid1.colColuna.Add ("Lcto. Bco.")
    objGrid1.colColuna.Add ("Documento")

    
   'campos de edição do grid
    objGrid1.colCampo.Add (Selecionado.Name)
    If Todos.Value = True Then objGrid1.colCampo.Add (Conciliado.Name)
    objGrid1.colCampo.Add (Data.Name)
    objGrid1.colCampo.Add (Valor.Name)
    objGrid1.colCampo.Add (Historico.Name)
    objGrid1.colCampo.Add (Categoria.Name)
    objGrid1.colCampo.Add (CodLctoBanco.Name)
    objGrid1.colCampo.Add (Documento.Name)
    
    If Todos.Value = True Then
        iGridDataCol = 3
        iGridValorCol = 4
        iGridHistoricoCol = 5
        iGridCategoriaCol = 6
        iGridCodLctoBancoCol = 7
        iGridDocumentoCol = 8
        iGridNumRefExternaCol = 6
    Else
        iGridDataCol = 2
        iGridValorCol = 3
        iGridHistoricoCol = 4
        iGridCategoriaCol = 5
        iGridCodLctoBancoCol = 6
        iGridDocumentoCol = 7
        iGridNumRefExternaCol = 5
    End If
    
    objGrid1.objGrid = GridExtrato
   
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 4
    
    objGrid1.objGrid.ColWidth(0) = 300
    
    objGrid1.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid1.iIncluirHScroll = GRID_INCLUIR_HSCROLL
        
    Call Grid_Inicializa(objGrid1)
    
    GridExtrato.HighLight = flexHighlightAlways
    
    Inicializa_GridExtrato = SUCESSO
    
End Function

Private Function Inicializa_GridMov1() As Long
   
    iExibeTodosMov1 = Todos.Value
   
    Set objGrid2 = New AdmGrid
   
    'tela em questão
    Set objGrid2.objForm = Me
    
    objGrid2.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid2.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'titulos do grid
    objGrid2.colColuna.Add ("")
    objGrid2.colColuna.Add ("Marcado")
    If Todos.Value = True Then objGrid2.colColuna.Add ("Conciliado")
    objGrid2.colColuna.Add ("Data")
    objGrid2.colColuna.Add ("Valor")
    objGrid2.colColuna.Add ("Histórico")
    objGrid2.colColuna.Add ("Ref.Externa")

   'campos de edição do grid
    objGrid2.colCampo.Add (SelecionadoMov1.Name)
    If Todos.Value = True Then objGrid2.colCampo.Add (ConciliadoMov1.Name)
    objGrid2.colCampo.Add (DataMov1.Name)
    objGrid2.colCampo.Add (ValorMov1.Name)
    objGrid2.colCampo.Add (HistoricoMov1.Name)
    objGrid2.colCampo.Add (NumRefExterna1.Name)
    
    If Todos.Value = True Then
        iGridDataCol = 3
        iGridValorCol = 4
        iGridHistoricoCol = 5
        iGridNumRefExternaCol = 6
    Else
        iGridDataCol = 2
        iGridValorCol = 3
        iGridHistoricoCol = 4
        iGridNumRefExternaCol = 5
    End If
    
    objGrid2.objGrid = GridMov1
   
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid2.iLinhasVisiveis = 11
    
    objGrid2.objGrid.ColWidth(0) = 300
    
    objGrid2.iGridLargAuto = GRID_LARGURA_MANUAL
        
    objGrid2.iIncluirHScroll = GRID_INCLUIR_HSCROLL
        
    Call Grid_Inicializa(objGrid2)
    
    Inicializa_GridMov1 = SUCESSO
    
End Function

Private Function Inicializa_GridMov() As Long
   
    Set objGrid3 = New AdmGrid
   
    'tela em questão
    Set objGrid3.objForm = Me
    
    objGrid3.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid3.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'titulos do grid
    objGrid3.colColuna.Add ("")
    objGrid3.colColuna.Add ("Marcado")
    If Todos.Value = True Then objGrid3.colColuna.Add ("Conciliado")
    objGrid3.colColuna.Add ("Data")
    objGrid3.colColuna.Add ("Valor")
    objGrid3.colColuna.Add ("Histórico")
    objGrid3.colColuna.Add ("Ref.Externa")

   'campos de edição do grid
    objGrid3.colCampo.Add (SelecionadoMov.Name)
    If Todos.Value = True Then objGrid3.colCampo.Add (ConciliadoMov.Name)
    objGrid3.colCampo.Add (DataMov.Name)
    objGrid3.colCampo.Add (ValorMov.Name)
    objGrid3.colCampo.Add (HistoricoMov.Name)
    objGrid3.colCampo.Add (NumRefExterna.Name)
    
    objGrid3.objGrid = GridMov
   
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid3.iLinhasVisiveis = 4
    
    objGrid3.objGrid.ColWidth(0) = 300
    
    objGrid3.iGridLargAuto = GRID_LARGURA_MANUAL
        
    objGrid3.iIncluirHScroll = GRID_INCLUIR_HSCROLL
        
    Call Grid_Inicializa(objGrid3)
        
    GridMov.HighLight = flexHighlightAlways
        
    Inicializa_GridMov = SUCESSO
    
End Function

Function Preenche_GridMov(colMovCCI As Collection) As Long
'preenche o grid com os movimentos de conta corrente passados na coleção colCCI

Dim lErro As Long
Dim iIndice As Integer
Dim objMovCCI As ClassMovContaCorrente

On Error GoTo Erro_Preenche_GridMov

    GridMov.Clear

    If colMovCCI.Count < objGrid3.iLinhasVisiveis Then
        objGrid3.objGrid.Rows = objGrid3.iLinhasVisiveis + 1
    Else
        objGrid3.objGrid.Rows = colMovCCI.Count + 2
    End If
    
    lErro = Inicializa_GridMov()
    If lErro <> SUCESSO Then Error 10905

    objGrid3.iLinhasExistentes = colMovCCI.Count

    'preenche o grid com os dados retornados na coleção colCCI
    For iIndice = 1 To colMovCCI.Count

        Set objMovCCI = colMovCCI.Item(iIndice)
        
        If Todos.Value = True Then
        
            If objMovCCI.iConciliado = NAO_CONCILIADO Then
                GridMov.TextMatrix(iIndice, GRID_CONCILIADO_COL) = S_NAO_CONCILIADO
            Else
                GridMov.TextMatrix(iIndice, GRID_CONCILIADO_COL) = S_CONCILIADO
            End If
            
        End If
        
        GridMov.TextMatrix(iIndice, iGridDataCol) = Format(objMovCCI.dtDataMovimento, "dd/mm/yyyy")
        GridMov.TextMatrix(iIndice, iGridValorCol) = Format(objMovCCI.dValor, "Standard")
        GridMov.TextMatrix(iIndice, iGridHistoricoCol) = objMovCCI.sHistorico
        GridMov.TextMatrix(iIndice, iGridNumRefExternaCol) = objMovCCI.sNumRefExterna
        
    Next

    Call Grid_Refresh_Checkbox(objGrid3)

    Call Soma_Coluna_Grid(objGrid3, iGridValorCol, Total3, False, GRID_SELECIONADO_COL)

    Preenche_GridMov = SUCESSO

    Exit Function

Erro_Preenche_GridMov:

    Preenche_GridMov = Err

    Select Case Err

        Case 10905

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154524)

    End Select

    Exit Function

End Function

Function Preenche_GridMov1(colMovCCI As Collection) As Long
'preenche o grid com os movimentos de conta corrente passados na coleção colCCI

Dim lErro As Long
Dim iIndice As Integer
Dim objMovCCI As ClassMovContaCorrente
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Preenche_GridMov1

    GridMov1.Clear

    If colMovCCI.Count < objGrid2.iLinhasVisiveis Then
        objGrid2.objGrid.Rows = objGrid2.iLinhasVisiveis + 1
    Else
        objGrid2.objGrid.Rows = colMovCCI.Count + 2
    End If

    lErro = Inicializa_GridMov1()
    If lErro <> SUCESSO Then Error 10906

    objGrid2.iLinhasExistentes = colMovCCI.Count

    'preenche o grid com os dados retornados na coleção colCCI
    For iIndice = 1 To colMovCCI.Count

        Set objMovCCI = colMovCCI.Item(iIndice)
        
        If Todos.Value = True Then
    
            If objMovCCI.iConciliado = NAO_CONCILIADO Then
                GridMov1.TextMatrix(iIndice, GRID_CONCILIADO_COL) = S_NAO_CONCILIADO
            Else
                GridMov1.TextMatrix(iIndice, GRID_CONCILIADO_COL) = S_CONCILIADO
            End If
        
        End If
        
        GridMov1.TextMatrix(iIndice, iGridDataCol) = Format(objMovCCI.dtDataMovimento, "dd/mm/yyyy")
        GridMov1.TextMatrix(iIndice, iGridValorCol) = Format(objMovCCI.dValor, "Standard")
        
        If Len(Trim(objMovCCI.sHistorico)) <> 0 Then
            GridMov1.TextMatrix(iIndice, iGridHistoricoCol) = objMovCCI.sHistorico
        Else
            For Each objCodigoNome In gcolTiposMovCta
                If objCodigoNome.iCodigo = objMovCCI.iTipo Then
                    GridMov1.TextMatrix(iIndice, iGridHistoricoCol) = objCodigoNome.sNome
                    Exit For
                End If
            Next
            
        End If
        GridMov1.TextMatrix(iIndice, iGridNumRefExternaCol) = objMovCCI.sNumRefExterna
        
    Next

    Call Grid_Refresh_Checkbox(objGrid2)
    
    Call Soma_Coluna_Grid(objGrid2, iGridValorCol, Total2, False, GRID_SELECIONADO_COL)

    Preenche_GridMov1 = SUCESSO

    Exit Function

Erro_Preenche_GridMov1:

    Preenche_GridMov1 = Err

    Select Case Err

        Case 10906

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154525)

    End Select

    Exit Function

End Function

Function Preenche_GridExtrato(colExtrato As Collection) As Long
'preenche o grid com os lançamento de extrato de conta corrente passados na coleção colExtrato

Dim lErro As Long
Dim iIndice As Integer
Dim objExtrBcoDet As ClassExtrBcoDet

On Error GoTo Erro_Preenche_GridExtrato

    GridExtrato.Clear

    If colExtrato.Count < objGrid1.iLinhasVisiveis Then
        objGrid1.objGrid.Rows = objGrid1.iLinhasVisiveis + 1
    Else
        objGrid1.objGrid.Rows = colExtrato.Count + 2
    End If

    lErro = Inicializa_GridExtrato()
    If lErro <> SUCESSO Then Error 10907

    objGrid1.iLinhasExistentes = colExtrato.Count

    'preenche o grid com os dados retornados na coleção colExtrato
    For iIndice = 1 To colExtrato.Count

        Set objExtrBcoDet = colExtrato.Item(iIndice)
        
        If Todos.Value = True Then
        
            If objExtrBcoDet.iConciliado = NAO_CONCILIADO Then
                GridExtrato.TextMatrix(iIndice, GRID_CONCILIADO_COL) = S_NAO_CONCILIADO
            Else
                GridExtrato.TextMatrix(iIndice, GRID_CONCILIADO_COL) = S_CONCILIADO
            End If
        
        End If
        
        GridExtrato.TextMatrix(iIndice, iGridDataCol) = Format(objExtrBcoDet.dtData, "dd/mm/yyyy")
        GridExtrato.TextMatrix(iIndice, iGridValorCol) = Format(objExtrBcoDet.dValor, "Standard")
        GridExtrato.TextMatrix(iIndice, iGridHistoricoCol) = objExtrBcoDet.sHistorico
        GridExtrato.TextMatrix(iIndice, iGridCategoriaCol) = CStr(objExtrBcoDet.iCategoria)
        GridExtrato.TextMatrix(iIndice, iGridCodLctoBancoCol) = objExtrBcoDet.sCodLctoBco
        GridExtrato.TextMatrix(iIndice, iGridDocumentoCol) = objExtrBcoDet.sDocumento
        
    Next

    Call Grid_Refresh_Checkbox(objGrid1)
    
    Call Soma_Coluna_Grid(objGrid1, iGridValorCol, Total1, False, GRID_SELECIONADO_COL)
    
    Preenche_GridExtrato = SUCESSO

    Exit Function

Erro_Preenche_GridExtrato:

    Preenche_GridExtrato = Err

    Select Case Err

        Case 10907

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154526)

    End Select

    Exit Function

End Function

Private Sub CodCCI_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_CodCCI_Validate

    If Len(Trim(CodCCI.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox CodCCI
    If CodCCI.Text = CodCCI.List(CodCCI.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(CodCCI, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 10736

    If lErro = 6730 Then
    
        objContaCorrenteInt.iCodigo = iCodigo
        'Lê a Conta Corrente Interna
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 10737
    
        'Se não encontrou a Conta Corrente Interna --> Erro
        If lErro = 11807 Then Error 10738
        
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            'Se a Conta não é da Filial selecionada
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43607

        End If
        
        CodCCI.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
    
    End If

    If lErro = 6731 Then Error 10739

    Exit Sub

Erro_CodCCI_Validate:

    Cancel = True


    Select Case Err

        Case 10736, 10737
        
        Case 10738
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)
        
            If vbMsgRes = vbYes Then
                'Lembrar de manter na tela o numero passado como parametro
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If
        
        Case 10739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, CodCCI.Text)
        
        Case 43607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, CodCCI.Text, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154527)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataDe.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then Error 10740
        
    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True


    Select Case Err

        Case 10740

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154528)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataAte.ClipText)) > 0 Then

        'verifica se a data final é válida
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then Error 10741
        
    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True


    Select Case Err

        Case 10741

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154529)

    End Select

    Exit Sub

End Sub

Private Sub Ordenados_Click()

Dim sOrdenacaoMov As String
Dim sOrdenacaoExt As String
Dim sSelecaoMov As String
Dim sSelecaoExt As String
Dim avCampo(1 To 6) As Variant
Dim iNumCampo As Integer
Dim lErro As Long
Dim colMovCCIAux As New Collection
Dim colExtratoAux As New Collection

On Error GoTo Erro_Ordenados_Click

    lErro = Monta_Ordenacao(sOrdenacaoMov, sOrdenacaoExt, Ordenados)
    If lErro <> SUCESSO Then Error 10795
    
    If sSQL_CNAB = "" Then Exit Sub
    
    lErro = Monta_Selecao(sSelecaoMov, sSelecaoExt, avCampo(), iNumCampo)
    If lErro <> SUCESSO Then Error 10908
    
    lErro = CF("MovContaCorrente_Le_Conciliacao", sSelecaoMov, avCampo(), iNumCampo, colMovCCIAux, sOrdenacaoMov)
    If lErro <> SUCESSO And lErro <> 10757 And lErro <> 10758 Then Error 10796
    
    'se não existirem lançamentos de extrato para a seleção em questão
    If lErro = 10757 Then Error 10797
        
    If lErro = 10758 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_MOV_ULTRAPASSOU_LIMITE")
    
    lErro = Gera_HistoricoMovto(colMovCCIAux)
    If lErro <> SUCESSO Then gError 96061
    
    lErro = CF("LctosExtratoBancario_Le_Conciliacao", sSelecaoExt, avCampo(), iNumCampo, colExtratoAux, sOrdenacaoExt)
    If lErro <> SUCESSO And lErro <> 10757 And lErro <> 10758 Then Error 10798
    
    'se não existirem lançamentos de extrato para a seleção em questão
    If lErro = 10768 Then Error 10799
        
    If lErro = 10769 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_LANCEXTRATO_ULTRAPASSOU_LIMITE")
    
    lErro = Preenche_GridMov(colMovCCIAux)
    If lErro <> SUCESSO Then Error 10800
    
    Set colMovCCI = colMovCCIAux
    
    lErro = Preenche_GridExtrato(colExtratoAux)
    If lErro <> SUCESSO Then Error 10801
    
    Set colExtrato = colExtratoAux
    
    Exit Sub

Erro_Ordenados_Click:

    Select Case Err

        Case 10795, 10796, 10798, 10800, 10801, 10908
    
        Case 10797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOS_INEXISTENTES_CONCILIACAO", Err)
    
        Case 10799
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCEXTRATO_INEXISTENTES_CONCILIACAO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154530)

    End Select

    Exit Sub

End Sub

Private Sub Ordenados1_Click()

Dim sOrdenacaoMov As String
Dim sOrdenacaoExt As String
Dim sSelecaoMov As String
Dim sSelecaoExt As String
Dim avCampo(1 To 6) As Variant
Dim iNumCampo As Integer
Dim colMovCCIAux As New Collection
Dim lErro As Long

On Error GoTo Erro_Ordenados1_Click

    lErro = Monta_Ordenacao(sOrdenacaoMov, sOrdenacaoExt, Ordenados1)
    If lErro <> SUCESSO Then Error 10802
    
    If sSQL_Papel = "" Then Exit Sub
    
    lErro = Monta_Selecao(sSelecaoMov, sSelecaoExt, avCampo(), iNumCampo)
    If lErro <> SUCESSO Then Error 10909
    
    lErro = CF("MovContaCorrente_Le_Conciliacao", sSelecaoMov, avCampo(), iNumCampo, colMovCCIAux, sOrdenacaoMov)
    If lErro <> SUCESSO And lErro <> 10757 And lErro <> 10758 Then Error 10803
    
    'se não existirem lançamentos de extrato para a seleção em questão
    If lErro = 10757 Then Error 10804
        
    If lErro = 10758 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_MOV_ULTRAPASSOU_LIMITE")
    
    lErro = Gera_HistoricoMovto(colMovCCIAux)
    If lErro <> SUCESSO Then gError 96061
    
    lErro = Preenche_GridMov1(colMovCCIAux)
    If lErro <> SUCESSO Then Error 10805
    
    Set colMovCCI1 = colMovCCIAux
    
    Exit Sub

Erro_Ordenados1_Click:

    Select Case Err

        Case 10802, 10803, 10805, 10909
    
        Case 10804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOS_INEXISTENTES_CONCILIACAO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154531)

    End Select

    Exit Sub

End Sub

Private Sub Selecionado_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call Soma_Coluna_Grid(objGrid1, iGridValorCol, Total1, False, GRID_SELECIONADO_COL)

End Sub

Private Sub Selecionado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Selecionado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SelecionadoMov_Click()

    iAlterado = REGISTRO_ALTERADO

    Call Soma_Coluna_Grid(objGrid3, iGridValorCol, Total3, False, GRID_SELECIONADO_COL)

End Sub

Private Sub SelecionadoMov_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid3)

End Sub

Private Sub SelecionadoMov_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid3)

End Sub

Private Sub SelecionadoMov_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid3.objControle = SelecionadoMov
    lErro = Grid_Campo_Libera_Foco(objGrid3)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SelecionadoMov1_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call Soma_Coluna_Grid(objGrid2, iGridValorCol, Total2, False, GRID_SELECIONADO_COL)

End Sub

Private Sub SelecionadoMov1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid2)

End Sub

Private Sub SelecionadoMov1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid2)

End Sub

Private Sub SelecionadoMov1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid2.objControle = SelecionadoMov1
    lErro = Grid_Campo_Libera_Foco(objGrid2)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorDe_Validate

    'Verifica se há um valor digitado
    If Len(Trim(ValorDe.Text)) > 0 Then
    
        'Critiva o valor digitado
        lErro = Valor_Critica(ValorDe.Text)
        If lErro <> SUCESSO Then Error 10742

        ValorDe.Text = Format(ValorDe.Text, "Fixed")

    End If

    Exit Sub

Erro_ValorDe_Validate:

    Cancel = True


    Select Case Err

        Case 10742

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154532)

    End Select

    Exit Sub

End Sub

Private Sub ValorAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorAte_Validate

    'Verifica se há um valor digitado
    If Len(Trim(ValorAte.Text)) > 0 Then
    
        'Critiva o valor digitado
        lErro = Valor_Critica(ValorAte.Text)
        If lErro <> SUCESSO Then Error 10743

        ValorAte.Text = Format(ValorAte.Text, "Fixed")

    End If

    Exit Sub

Erro_ValorAte_Validate:

    Cancel = True


    Select Case Err

        Case 10743

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154533)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    DataDe.SetFocus

    If Len(DataDe.ClipText) > 0 Then

        sData = DataDe.Text
        
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 10782
        
        DataDe.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_DownClick:
    
    Select Case Err
    
        Case 10782
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154534)
        
    End Select
    
    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    DataDe.SetFocus

    If Len(DataDe.ClipText) > 0 Then

        sData = DataDe.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 10783
        
        DataDe.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_UpClick:
    
    Select Case Err
    
        Case 10783
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154535)
        
    End Select
    
    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown2_DownClick

    DataAte.SetFocus

    If Len(DataAte.ClipText) > 0 Then

        sData = DataAte.Text
        
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 10784
        
        DataAte.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown2_DownClick:
    
    Select Case Err
    
        Case 10784
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154536)
        
    End Select
    
    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown2_UpClick

    DataAte.SetFocus

    If Len(DataAte.ClipText) > 0 Then

        sData = DataAte.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 10785
        
        DataAte.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown2_UpClick:
    
    Select Case Err
    
        Case 10785
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154537)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Conciliado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Conciliado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Conciliado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Conciliado
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ConciliadoMov_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid3)

End Sub

Private Sub ConciliadoMov_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid3)

End Sub

Private Sub ConciliadoMov_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid3.objControle = ConciliadoMov
    lErro = Grid_Campo_Libera_Foco(objGrid3)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ConciliadoMov1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid2)

End Sub

Private Sub ConciliadoMov1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid2)

End Sub

Private Sub ConciliadoMov1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid2.objControle = ConciliadoMov1
    lErro = Grid_Campo_Libera_Foco(objGrid2)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridExtrato_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridExtrato_GotFocus()
    Call Grid_Recebe_Foco(objGrid1)
End Sub

Private Sub GridExtrato_EnterCell()
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
End Sub

Private Sub GridExtrato_LeaveCell()
    Call Saida_Celula(objGrid1)
End Sub

Private Sub GridExtrato_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
End Sub

Private Sub GridExtrato_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridExtrato_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGrid1)
End Sub

Private Sub GridExtrato_RowColChange()
    Call Grid_RowColChange(objGrid1)
End Sub

Private Sub GridExtrato_Scroll()
    Call Grid_Scroll(objGrid1)
End Sub

Private Sub GridMov_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid3, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid3, iAlterado)
    End If

End Sub

Private Sub GridMov_GotFocus()
    Call Grid_Recebe_Foco(objGrid3)
End Sub

Private Sub GridMov_EnterCell()
    Call Grid_Entrada_Celula(objGrid3, iAlterado)
End Sub

Private Sub GridMov_LeaveCell()
    Call Saida_Celula(objGrid3)
End Sub

Private Sub GridMov_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGrid3)
End Sub

Private Sub GridMov_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid3, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid3, iAlterado)
    End If

End Sub

Private Sub GridMov_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGrid3)
End Sub

Private Sub GridMov_RowColChange()
    Call Grid_RowColChange(objGrid3)
End Sub

Private Sub GridMov_Scroll()
    Call Grid_Scroll(objGrid3)
End Sub


Private Sub GridMov1_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid2, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid2, iAlterado)
    End If

End Sub

Private Sub GridMov1_GotFocus()
    Call Grid_Recebe_Foco(objGrid2)
End Sub

Private Sub GridMov1_EnterCell()
    Call Grid_Entrada_Celula(objGrid2, iAlterado)
End Sub

Private Sub GridMov1_LeaveCell()
    Call Saida_Celula(objGrid2)
End Sub

Private Sub GridMov1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGrid2)
End Sub

Private Sub GridMov1_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid2, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid2, iAlterado)
    End If

End Sub

Private Sub GridMov1_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGrid2)
End Sub

Private Sub GridMov1_RowColChange()
    Call Grid_RowColChange(objGrid2)
End Sub

Private Sub GridMov1_Scroll()
    Call Grid_Scroll(objGrid2)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente /m

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case GRID_SELECIONADO_COL
                lErro = Saida_Celula_Selecionado(objGridInt)
                If lErro <> SUCESSO Then Error 10786
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 10787

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 10786

        Case 10787
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154538)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Selecionado(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Selecionado

    If objGridInt.objGrid.Name = GridExtrato.Name Then Set objGridInt.objControle = Selecionado

    If objGridInt.objGrid.Name = GridMov1.Name Then Set objGridInt.objControle = SelecionadoMov1

    If objGridInt.objGrid.Name = GridMov.Name Then Set objGridInt.objControle = SelecionadoMov

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 10788

    Saida_Celula_Selecionado = SUCESSO

    Exit Function

Erro_Saida_Celula_Selecionado:

    Saida_Celula_Selecionado = Err

    Select Case Err

        Case 10788
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154539)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONCILIACAO_BANCARIA_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Conciliação Bancária"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConciliacaoBancaria"
    
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

'***** fim do trecho a ser copiado ******




Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim lErroAux As Long
Dim objMovCtaCorrente As ClassMovContaCorrente

On Error GoTo Erro_Botao_ProcurarMov_Click

    'verifica se alguma linha está selecionada
    If GridMov1.Row < 1 Then gError 189365
    
    Set objMovCtaCorrente = colMovCCI1.Item(GridMov1.Row)
    
    lErro = CF("MovCtaCorrenteLista_BotaoEdita", objMovCtaCorrente, lErroAux)
    If lErro <> SUCESSO Then gError 189366
    
    Exit Sub
    
Erro_Botao_ProcurarMov_Click:

    Select Case gErr
    
        Case 189365 'ERRO_LINHA_GRID_NAO_SELECIONADA
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 189366

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189367)
            
    End Select
    
    Exit Sub

End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVGeracaoNF 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5025
      Index           =   1
      Left            =   210
      TabIndex        =   25
      Top             =   825
      Width           =   9120
      Begin VB.CheckBox GerarNFParaCadaFat 
         Caption         =   "Gerar uma NF para cada Fatura"
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
         Left            =   5265
         TabIndex        =   2
         Top             =   105
         Width           =   3120
      End
      Begin VB.CheckBox optExtraiDados 
         Caption         =   "Extrair dados sobre o passageiro no Sigav mesmo se a informação já existir no Corporator"
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
         Left            =   735
         TabIndex        =   4
         Top             =   450
         Visible         =   0   'False
         Width           =   8055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   4620
         Left            =   165
         TabIndex        =   29
         Top             =   405
         Width           =   8745
         Begin VB.ComboBox Marca 
            Height          =   315
            ItemData        =   "TRVGeracaoNF.ctx":0000
            Left            =   3075
            List            =   "TRVGeracaoNF.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   300
            Width           =   1635
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data de baixa dos títulos"
            Height          =   1275
            Left            =   6435
            TabIndex        =   66
            Top             =   3285
            Width           =   2190
            Begin MSComCtl2.UpDown UpDownDataBaixaDe 
               Height          =   300
               Left            =   1845
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   330
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataBaixaDe 
               Height          =   300
               Left            =   675
               TabIndex        =   19
               Top             =   315
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataBaixaAte 
               Height          =   300
               Left            =   675
               TabIndex        =   21
               Top             =   825
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataBaixaAte 
               Height          =   300
               Left            =   1845
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   840
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label4 
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
               Left            =   255
               TabIndex        =   68
               Top             =   885
               Width           =   360
            End
            Begin VB.Label Label3 
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
               Left            =   300
               TabIndex        =   67
               Top             =   375
               Width           =   315
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Títulos"
            Height          =   1170
            Index           =   5
            Left            =   6435
            TabIndex        =   63
            Top             =   720
            Width           =   2190
            Begin MSMask.MaskEdBox FaturaDe 
               Height          =   300
               Left            =   720
               TabIndex        =   10
               Top             =   255
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox FaturaAte 
               Height          =   300
               Left            =   720
               TabIndex        =   11
               Top             =   675
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin VB.Label LabelFim 
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
               Left            =   285
               TabIndex        =   65
               Top             =   735
               Width           =   360
            End
            Begin VB.Label LabelInicio 
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
               Left            =   330
               TabIndex        =   64
               Top             =   285
               Width           =   315
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Tipo de Documento"
            Height          =   1170
            Index           =   0
            Left            =   150
            TabIndex        =   62
            Top             =   720
            Width           =   4530
            Begin VB.OptionButton TipoDocApenas 
               Caption         =   "Apenas:"
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
               Left            =   405
               TabIndex        =   6
               Top             =   675
               Width           =   1050
            End
            Begin VB.OptionButton TipoDocTodos 
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
               Left            =   405
               TabIndex        =   5
               Top             =   315
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.ComboBox TipoDocSeleciona 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "TRVGeracaoNF.ctx":0033
               Left            =   1470
               List            =   "TRVGeracaoNF.ctx":0035
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   645
               Width           =   2955
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data de emissão dos títulos"
            Height          =   1275
            Left            =   6435
            TabIndex        =   33
            Top             =   1965
            Width           =   2190
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1845
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   330
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   675
               TabIndex        =   15
               Top             =   315
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   675
               TabIndex        =   17
               Top             =   825
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   1845
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   840
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label2 
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
               Left            =   300
               TabIndex        =   35
               Top             =   375
               Width           =   315
            End
            Begin VB.Label Label5 
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
               Left            =   255
               TabIndex        =   34
               Top             =   885
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Filiais"
            Height          =   2595
            Left            =   150
            TabIndex        =   32
            Top             =   1965
            Width           =   6180
            Begin MSMask.MaskEdBox FEPerc 
               Height          =   255
               Left            =   3420
               TabIndex        =   71
               Top             =   1050
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   450
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   8
               Format          =   "0%"
               PromptChar      =   " "
            End
            Begin VB.CheckBox FESel 
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
               Left            =   690
               TabIndex        =   70
               Top             =   1050
               Width           =   750
            End
            Begin VB.CommandButton BotaoMarcarTodos 
               Caption         =   "Marcar Todos"
               Height          =   540
               Left            =   4605
               Picture         =   "TRVGeracaoNF.ctx":0037
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   240
               Width           =   1440
            End
            Begin VB.CommandButton BotaoDesmarcarTodos 
               Caption         =   "Desmarcar Todos"
               Height          =   540
               Left            =   4620
               Picture         =   "TRVGeracaoNF.ctx":1051
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   885
               Width           =   1440
            End
            Begin MSMask.MaskEdBox FE 
               Height          =   255
               Left            =   1485
               TabIndex        =   69
               Top             =   1035
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   450
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
            Begin MSFlexGridLib.MSFlexGrid GridFE 
               Height          =   405
               Left            =   120
               TabIndex        =   12
               Top             =   225
               Width           =   4410
               _ExtentX        =   7779
               _ExtentY        =   714
               _Version        =   393216
               Rows            =   16
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               FocusRect       =   2
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Títulos"
            Height          =   1185
            Left            =   4785
            TabIndex        =   30
            Top             =   705
            Width           =   1545
            Begin VB.OptionButton OptBaixado 
               Caption         =   "Baixados"
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
               Left            =   285
               TabIndex        =   9
               Top             =   705
               Width           =   1110
            End
            Begin VB.OptionButton OptEmitido 
               Caption         =   "Emitidos"
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
               Left            =   300
               TabIndex        =   8
               Top             =   330
               Value           =   -1  'True
               Width           =   1110
            End
         End
         Begin VB.Label Label8 
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
            Left            =   2430
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   81
            Top             =   345
            Width           =   600
         End
      End
      Begin MSComCtl2.UpDown UpDownDataEmissao 
         Height          =   300
         Left            =   4425
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   3255
         TabIndex        =   0
         Top             =   45
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data de Emissão das NFs:"
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
         Left            =   975
         TabIndex        =   80
         Top             =   105
         Width           =   2250
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5025
      Index           =   2
      Left            =   195
      TabIndex        =   26
      Top             =   825
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton BotaoTitSemNF 
         Caption         =   "Títulos por Filial"
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
         Index           =   2
         Left            =   6840
         TabIndex        =   59
         Top             =   4500
         Width           =   2145
      End
      Begin MSMask.MaskEdBox PercR 
         Height          =   240
         Left            =   5475
         TabIndex        =   73
         Top             =   1365
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin VB.TextBox ValorR 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   4080
         TabIndex        =   72
         Top             =   1380
         Width           =   1125
      End
      Begin VB.CommandButton BotaoGerar 
         Caption         =   "Gerar Notas Fiscais"
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
         Left            =   75
         TabIndex        =   58
         Top             =   4500
         Width           =   2145
      End
      Begin VB.TextBox NFAte 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   7950
         TabIndex        =   41
         Top             =   2670
         Width           =   795
      End
      Begin VB.TextBox NFDe 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   6780
         TabIndex        =   40
         Top             =   2655
         Width           =   1050
      End
      Begin VB.TextBox Valor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   5415
         TabIndex        =   39
         Top             =   2625
         Width           =   1125
      End
      Begin VB.TextBox NumTitulos 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   4260
         TabIndex        =   38
         Top             =   2625
         Width           =   660
      End
      Begin VB.TextBox NumNF 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3060
         TabIndex        =   37
         Top             =   2610
         Width           =   690
      End
      Begin VB.TextBox FilialEmpresa 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1260
         TabIndex        =   31
         Top             =   1365
         Width           =   1755
      End
      Begin MSFlexGridLib.MSFlexGrid GridFilial 
         Height          =   660
         Left            =   60
         TabIndex        =   27
         Top             =   45
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   1164
         _Version        =   393216
         Rows            =   11
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   6930
         TabIndex        =   61
         Top             =   4065
         Width           =   1620
      End
      Begin VB.Label Label1 
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
         Height          =   300
         Left            =   5865
         TabIndex        =   60
         Top             =   4140
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5025
      Index           =   3
      Left            =   195
      TabIndex        =   42
      Top             =   825
      Visible         =   0   'False
      Width           =   9030
      Begin VB.CommandButton BotaoTitSemNF 
         Caption         =   "Títulos por Cliente"
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
         Index           =   3
         Left            =   6840
         TabIndex        =   75
         Top             =   4500
         Width           =   2145
      End
      Begin VB.TextBox NFFilialEmpresaR 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   0
         TabIndex        =   74
         Top             =   0
         Width           =   1785
      End
      Begin VB.TextBox NFFilialEmpresa 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   600
         TabIndex        =   56
         Top             =   3690
         Width           =   1785
      End
      Begin VB.TextBox NFNumItens 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   5910
         TabIndex        =   46
         Top             =   3225
         Width           =   435
      End
      Begin VB.TextBox NFValor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   7260
         TabIndex        =   45
         Top             =   3225
         Width           =   915
      End
      Begin VB.TextBox NFCliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   2985
         TabIndex        =   44
         Top             =   3240
         Width           =   2340
      End
      Begin VB.TextBox NFNumNota 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1935
         TabIndex        =   43
         Top             =   3225
         Width           =   780
      End
      Begin MSFlexGridLib.MSFlexGrid GridNF 
         Height          =   1485
         Left            =   60
         TabIndex        =   47
         Top             =   45
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   2619
         _Version        =   393216
         Rows            =   11
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5025
      Index           =   4
      Left            =   195
      TabIndex        =   48
      Top             =   825
      Visible         =   0   'False
      Width           =   9015
      Begin VB.TextBox ItemFat 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1365
         TabIndex        =   79
         Top             =   3735
         Width           =   705
      End
      Begin VB.CommandButton BotaoConsTitulo 
         Caption         =   "Consultar Título"
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
         Left            =   6840
         TabIndex        =   76
         Top             =   4500
         Width           =   2145
      End
      Begin VB.TextBox ItemFilialEmpresa 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   300
         TabIndex        =   57
         Top             =   3450
         Width           =   1755
      End
      Begin VB.TextBox ItemValor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   7545
         TabIndex        =   54
         Top             =   3450
         Width           =   975
      End
      Begin VB.TextBox ItemDesc 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   6165
         TabIndex        =   53
         Top             =   3450
         Width           =   870
      End
      Begin VB.TextBox ItemProduto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   5400
         TabIndex        =   52
         Top             =   3450
         Width           =   510
      End
      Begin VB.TextBox ItemCliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3585
         TabIndex        =   51
         Top             =   3450
         Width           =   1785
      End
      Begin VB.TextBox Item 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3090
         TabIndex        =   50
         Top             =   3450
         Width           =   480
      End
      Begin VB.TextBox ItemNumNota 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   2115
         TabIndex        =   49
         Top             =   3450
         Width           =   840
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   3195
         Left            =   60
         TabIndex        =   55
         Top             =   45
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   5636
         _Version        =   393216
         Rows            =   11
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label DescItem 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1530
         TabIndex        =   78
         Top             =   4500
         Width           =   5175
      End
      Begin VB.Label Label6 
         Caption         =   "Descrição Item:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   77
         Top             =   4515
         Width           =   1410
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   8220
      ScaleHeight     =   480
      ScaleWidth      =   1095
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   75
      Width           =   1155
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   570
         Picture         =   "TRVGeracaoNF.ctx":2233
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   60
         Picture         =   "TRVGeracaoNF.ctx":23B1
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStripOpcao 
      Height          =   5535
      Left            =   135
      TabIndex        =   28
      Top             =   405
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   9763
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filial Empresa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notas Fiscais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "TRVGeracaoNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer 'apenas p/uso pela interface c/grid
Dim iFrameAtual As Integer
Dim iFrameSelecaoAlterado As Integer

'Grid Filial
Dim objGridFilial As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_NumNF_Col As Integer
Dim iGrid_NumTitulos_Col As Integer
Dim iGrid_Valor_Col  As Integer
Dim iGrid_ValorR_Col  As Integer
Dim iGrid_PercR_Col  As Integer
Dim iGrid_NFDe_Col As Integer
Dim iGrid_NFAte_Col As Integer

'Grid NF
Dim objGridNF As AdmGrid
Dim iGrid_NFFilialEmpresa_Col As Integer
Dim iGrid_NFFilialEmpresaR_Col As Integer
Dim iGrid_NFNumNota_Col As Integer
Dim iGrid_NFCliente_Col As Integer
Dim iGrid_NFValor_Col  As Integer
Dim iGrid_NFNumItens_Col As Integer

'Grid Itens
Dim objGridItens As AdmGrid
Dim iGrid_ItemFilialEmpresa_Col As Integer
Dim iGrid_ItemNumNota_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_ItemCliente_Col As Integer
Dim iGrid_ItemProduto_Col  As Integer
Dim iGrid_ItemFat_Col  As Integer
Dim iGrid_ItemDesc_Col  As Integer
Dim iGrid_ItemValor_Col  As Integer

'Grid Itens
Dim objGridFE As AdmGrid
Dim iGrid_FESel_Col As Integer
Dim iGrid_FE_Col As Integer
Dim iGrid_FEPerc_Col As Integer

Dim gobjGeracaoNF As New ClassTRVGeracaoNF
Dim gcolFiliaisTela As Collection
Dim gcolNFOrd As Collection
Dim gcolTitOrd As Collection

'Eventos de Browse
Private WithEvents objEventoTituloDe As AdmEvento
Attribute objEventoTituloDe.VB_VarHelpID = -1
Private WithEvents objEventoTituloAte As AdmEvento
Attribute objEventoTituloAte.VB_VarHelpID = -1

'CONTANTES GLOBAIS DA TELA
Const TAB_Selecao = 1
Const TAB_GERACAO = 2
Const TAB_NF = 3
Const TAB_ITENS = 4

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridFilial = Nothing
    Set objGridNF = Nothing
    Set objGridFE = Nothing
    Set objGridItens = Nothing
    Set gobjGeracaoNF = Nothing
    Set gcolFiliaisTela = Nothing
    Set gcolNFOrd = Nothing
    Set gcolTitOrd = Nothing
    
End Sub

Private Sub DataAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Se a DataAte está preenchida
    If Len(DataAte.ClipText) > 0 Then

        'Verifica se a DataAte é válida
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 192240

    End If
    
    Exit Sub

Erro_DataAte_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 192240 'Tratado na rotina chamada

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192241)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Se a DataDe está preenchida
    If Len(DataDe.ClipText) > 0 Then

        'Verifica se a DataDe é válida
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 192242

    End If

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 192242

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192243)

    End Select

    Exit Sub

End Sub

Private Sub FiliaisEmpresa_Click()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GerarNFParaCadaFat_Click()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GridFilial_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFilial, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFilial, iAlterado)
    End If

End Sub

Private Sub GridFilial_GotFocus()
    Call Grid_Recebe_Foco(objGridFilial)
End Sub

Private Sub GridFilial_EnterCell()
    Call Grid_Entrada_Celula(objGridFilial, iAlterado)
End Sub

Private Sub GridFilial_LeaveCell()
    Call Saida_Celula(objGridFilial)
End Sub

Private Sub GridFilial_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridFilial)
End Sub

Private Sub GridFilial_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFilial, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFilial, iAlterado)
    End If

End Sub

Private Sub GridFilial_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridFilial)
End Sub

Private Sub GridFilial_RowColChange()
    Call Grid_RowColChange(objGridFilial)
End Sub

Private Sub GridFilial_Scroll()
    Call Grid_Scroll(objGridFilial)
End Sub

Private Sub GridNF_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If
    
    Call Ordenacao_ClickGrid(objGridNF)

End Sub

Private Sub GridNF_GotFocus()
    Call Grid_Recebe_Foco(objGridNF)
End Sub

Private Sub GridNF_EnterCell()
    Call Grid_Entrada_Celula(objGridNF, iAlterado)
End Sub

Private Sub GridNF_LeaveCell()
    Call Saida_Celula(objGridNF)
End Sub

Private Sub GridNF_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridNF)
End Sub

Private Sub GridNF_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If

End Sub

Private Sub GridNF_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridNF)
End Sub

Private Sub GridNF_RowColChange()
    Call Grid_RowColChange(objGridNF)
End Sub

Private Sub GridNF_Scroll()
    Call Grid_Scroll(objGridNF)
End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
    colcolColecoes.Add gcolTitOrd

    Call Ordenacao_ClickGrid(objGridItens, , colcolColecoes)

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

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
    If GridItens.Row <> 0 Then DescItem.Caption = GridItens.TextMatrix(GridItens.Row, iGrid_ItemDesc_Col)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub Form_Load()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objGridFilial = New AdmGrid
    Set objGridNF = New AdmGrid
    Set objGridItens = New AdmGrid
    Set objGridFE = New AdmGrid
    
    'Inicializa os Eventos de Browser
    Set objEventoTituloDe = New AdmEvento
    Set objEventoTituloAte = New AdmEvento
    
    'Executa a Inicialização do grid
    lErro = Inicializa_Grid_Filial(objGridFilial)
    If lErro <> SUCESSO Then gError 192244
    
    'Executa a Inicialização do grid
    lErro = Inicializa_Grid_NF(objGridNF)
    If lErro <> SUCESSO Then gError 192245
    
    'Executa a Inicialização do grid
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 192246
    
    'Executa a Inicialização do grid
    lErro = Inicializa_Grid_FE(objGridFE)
    If lErro <> SUCESSO Then gError 192246
    
    lErro = Carrega_FilialEmpresa
    If lErro <> SUCESSO Then gError 192247
    
    lErro = Carrega_TipoDocumento(TipoDocSeleciona)
    If lErro <> SUCESSO Then gError 192305
    
    Call Default_Tela
        
    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 192244, 192245, 192246, 192247, 192305
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192248)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Move_Selecao_Memoria(ByVal objGeracaoNF As ClassTRVGeracaoNF) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objSenha As New ClassSenha
Dim objNFFilial As ClassTRVGeracaoNFFiliais

On Error GoTo Erro_Move_Selecao_Memoria

    objGeracaoNF.dtDataDe = StrParaDate(DataDe.Text)
    objGeracaoNF.dtDataAte = StrParaDate(DataAte.Text)

    objGeracaoNF.dtDataBaixaDe = StrParaDate(DataBaixaDe.Text)
    objGeracaoNF.dtDataBaixaAte = StrParaDate(DataBaixaAte.Text)

    objGeracaoNF.lTituloDe = StrParaLong(FaturaDe.Text)
    objGeracaoNF.lTituloAte = StrParaLong(FaturaAte.Text)
    objGeracaoNF.iEmpresa = Marca.ItemData(Marca.ListIndex)
    
    If optExtraiDados.Value = vbChecked Then
        objGeracaoNF.iExtrairDados = MARCADO
    Else
        objGeracaoNF.iExtrairDados = DESMARCADO
    End If
    
    If StrParaDate(DataEmissao.Text) = DATA_NULA Then gError 198193
    
    objGeracaoNF.dtDataEmissao = StrParaDate(DataEmissao.Text)

    'Se DataDe e DataAté estão preenchidas
    If objGeracaoNF.dtDataDe <> DATA_NULA And objGeracaoNF.dtDataAte <> DATA_NULA Then
        'Verifica se DataAté é maior ou igual a DataDe
        If objGeracaoNF.dtDataAte < objGeracaoNF.dtDataDe Then gError 192249
    End If
    
    'Se DataDe e DataAté estão preenchidas
    If objGeracaoNF.dtDataBaixaDe <> DATA_NULA And objGeracaoNF.dtDataBaixaAte <> DATA_NULA Then
        'Verifica se DataAté é maior ou igual a DataDe
        If objGeracaoNF.dtDataBaixaAte < objGeracaoNF.dtDataBaixaDe Then gError 192249
    End If
    
    'Se DataDe e DataAté estão preenchidas
    If objGeracaoNF.lTituloDe <> 0 And objGeracaoNF.lTituloAte <> 0 Then
        'Verifica se DataAté é maior ou igual a DataDe
        If objGeracaoNF.lTituloAte < objGeracaoNF.lTituloDe Then gError 196450
    End If

    If OptBaixado.Value Then
        objGeracaoNF.iTipo = TRV_GERACAONF_TITULOS_BAIXADO
    Else
        objGeracaoNF.iTipo = TRV_GERACAONF_TITULOS_EMITIDOS
    End If
'
'    For iLinha = 0 To FiliaisEmpresa.ListCount - 1
'        If FiliaisEmpresa.Selected(iLinha) Then
'            objGeracaoNF.colFiliais.Add FiliaisEmpresa.ItemData(iLinha)
'        End If
'    Next
    
    For iLinha = 1 To objGridFE.iLinhasExistentes
        If StrParaInt(GridFE.TextMatrix(iLinha, iGrid_FESel_Col)) = MARCADO Then
        
            Set objNFFilial = New ClassTRVGeracaoNFFiliais
            
            objNFFilial.iFilialEmpresa = Codigo_Extrai(GridFE.TextMatrix(iLinha, iGrid_FE_Col))
            objNFFilial.dPercentual = PercentParaDbl(GridFE.TextMatrix(iLinha, iGrid_FEPerc_Col))
            
            objGeracaoNF.colFiliais.Add objNFFilial
            
        End If
    Next
    
    If TipoDocApenas.Value = True Then
        objGeracaoNF.sTipoDoc = SCodigo_Extrai(TipoDocSeleciona.Text)
    Else
        objGeracaoNF.sTipoDoc = ""
    End If
    
    If GerarNFParaCadaFat.Value = vbChecked Then
        objGeracaoNF.iGerarNFParaCadaFat = MARCADO
    Else
        objGeracaoNF.iGerarNFParaCadaFat = DESMARCADO
    End If
    
'    Load SigavSenha
'    lErro = SigavSenha.Trata_Parametros(objSenha)
'    If lErro <> SUCESSO Then gError 192728
'    SigavSenha.Show vbModal
'
'    If Len(Trim(objSenha.sSenha)) = 0 Then gError 192729
'
'    objGeracaoNF.sSenhaSigav = objSenha.sSenha

    Move_Selecao_Memoria = SUCESSO

    Exit Function
    
Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr
    
    Select Case gErr
        
        Case 192249
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", gErr)
            
        Case 192728
        
        Case 192729
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", gErr)
            
        Case 196450
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_MENOR_NUMERO_DE", gErr)
            
        Case 198193
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192250)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Marca_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OptBaixado_Click()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OptEmitido_Click()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TabStripOpcao_Click()

Dim lErro As Long

On Error GoTo Erro_TabStripOpcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStripOpcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStripOpcao, Me) <> SUCESSO Then Exit Sub
       
        'Torna Frame de Bloqueios visível
        Frame1(TabStripOpcao.SelectedItem.Index).Visible = True
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStripOpcao.SelectedItem.Index
       
        'Se o frame anterior foi o de Seleção e ele foi alterado
        If iFrameAtual <> TAB_Selecao And iFrameSelecaoAlterado = REGISTRO_ALTERADO Then
    
            DoEvents
    
            lErro = Traz_GeracaoNF_Tela
            If lErro <> SUCESSO Then gError 192251
    
            iFrameSelecaoAlterado = 0
    
        End If
    
    End If

    Exit Sub

Erro_TabStripOpcao_Click:

    Select Case gErr
    
        Case 192251

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192252)

    End Select

    Exit Sub

End Sub

Private Function Traz_GeracaoNF_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iItem As Integer
Dim objGeracaoNF As New ClassTRVGeracaoNF
Dim objGeracaoNFItem As ClassTRVGeracaoNFItem
Dim objCodigoNomeO As New AdmCodigoNome
Dim objCodigoNomeD As New AdmCodigoNome
Dim objNF As New ClassNFiscal
Dim objItemNF As New ClassItemNF
Dim objcliente As ClassCliente
Dim iTotalNF As Integer
Dim iTotalItens As Integer
Dim dValorTotal As Double
Dim sParte As String
Dim iPOS As Integer
Dim vObjeto As Variant

On Error GoTo Erro_Traz_GeracaoNF_Tela

    Call Grid_Limpa(objGridFilial)
    Call Grid_Limpa(objGridNF)
    Call Grid_Limpa(objGridItens)
    
    Call Ordenacao_Limpa(objGridFilial)
    Call Ordenacao_Limpa(objGridNF)
    Call Ordenacao_Limpa(objGridItens)
    
    lErro = Move_Selecao_Memoria(objGeracaoNF)
    If lErro <> SUCESSO Then gError 192253
  
    GL_objMDIForm.MousePointer = vbHourglass
  
    'Preenche a Coleção de Bloqueios
    lErro = CF("TRVGeracaoNF_Le_Dados", objGeracaoNF)
    If lErro <> SUCESSO Then gError 192254
    
    If objGeracaoNF.colTitulos.Count = 0 Then gError 192255
    
    Set gobjGeracaoNF = objGeracaoNF
    Set gcolNFOrd = New Collection
    Set gcolTitOrd = New Collection
    
    For Each vObjeto In gobjGeracaoNF.colNF
        gcolNFOrd.Add vObjeto
    Next
    
    For Each vObjeto In gobjGeracaoNF.colTitulos
        gcolTitOrd.Add vObjeto
    Next
    
    iIndice = 0
    For Each objGeracaoNFItem In objGeracaoNF.colItens
    
        'If objGeracaoNFItem.iNumNF <> 0 Then
    
            iIndice = iIndice + 1
            iTotalNF = iTotalNF + objGeracaoNFItem.iNumNF
            iTotalItens = iTotalItens + objGeracaoNFItem.iNumTitulos
            dValorTotal = dValorTotal + objGeracaoNFItem.dValor
            
            For Each objCodigoNomeO In gcolFiliaisTela
                If objGeracaoNFItem.iFilialEmpresa = objCodigoNomeO.iCodigo Then
                    objGeracaoNFItem.sFilialEmpresa = objCodigoNomeO.sNome
                    Exit For
                End If
            Next
        
            GridFilial.TextMatrix(iIndice, iGrid_FilialEmpresa_Col) = objGeracaoNFItem.iFilialEmpresa & SEPARADOR & objGeracaoNFItem.sFilialEmpresa
            
            If objGeracaoNFItem.iNumNF <> 0 Then
                GridFilial.TextMatrix(iIndice, iGrid_NumNF_Col) = CStr(objGeracaoNFItem.iNumNF)
            Else
                GridFilial.TextMatrix(iIndice, iGrid_NumNF_Col) = ""
            End If
            
            If objGeracaoNFItem.iNumTitulos <> 0 Then
                GridFilial.TextMatrix(iIndice, iGrid_NumTitulos_Col) = CStr(objGeracaoNFItem.iNumTitulos)
            Else
                GridFilial.TextMatrix(iIndice, iGrid_NumTitulos_Col) = ""
            End If
            
            If objGeracaoNFItem.dValor <> 0 Then
                GridFilial.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objGeracaoNFItem.dValor, "STANDARD")
            Else
                GridFilial.TextMatrix(iIndice, iGrid_Valor_Col) = ""
            End If
            
            If objGeracaoNFItem.dValorR <> 0 Then
                GridFilial.TextMatrix(iIndice, iGrid_ValorR_Col) = Format(objGeracaoNFItem.dValorR, "STANDARD")
            Else
                GridFilial.TextMatrix(iIndice, iGrid_ValorR_Col) = ""
            End If
            
            If objGeracaoNFItem.dValor <> 0 Then
                GridFilial.TextMatrix(iIndice, iGrid_PercR_Col) = Format(1 - (objGeracaoNFItem.dValorR / objGeracaoNFItem.dValor), "PERCENT")
            Else
                GridFilial.TextMatrix(iIndice, iGrid_PercR_Col) = ""
            End If
            
            If objGeracaoNFItem.lNFDe <> 0 Then
                GridFilial.TextMatrix(iIndice, iGrid_NFDe_Col) = CStr(objGeracaoNFItem.lNFDe)
            Else
                GridFilial.TextMatrix(iIndice, iGrid_NFDe_Col) = ""
            End If
            
            If objGeracaoNFItem.lNFAte <> 0 Then
                GridFilial.TextMatrix(iIndice, iGrid_NFAte_Col) = CStr(objGeracaoNFItem.lNFAte)
            Else
                GridFilial.TextMatrix(iIndice, iGrid_NFAte_Col) = ""
            End If
            
        'End If
        
    Next
    
    objGridFilial.iLinhasExistentes = iIndice
    
    ValorTotal.Caption = Format(dValorTotal, "STANDARD")
    
    If GridNF.Rows <= iTotalNF Then Call Refaz_Grid(objGridNF, iTotalNF)
    If GridItens.Rows <= iTotalItens Then Call Refaz_Grid(objGridItens, iTotalItens)
    
    iIndice = 0
    iItem = 0
    For Each objNF In objGeracaoNF.colNF
    
        Set objcliente = New ClassCliente
    
        iIndice = iIndice + 1
        
        For Each objCodigoNomeD In gcolFiliaisTela
            If objNF.iFilialEmpresa = objCodigoNomeD.iCodigo Then
                Exit For
            End If
        Next
    
        GridNF.TextMatrix(iIndice, iGrid_NFFilialEmpresaR_Col) = objCodigoNomeD.iCodigo & SEPARADOR & objCodigoNomeD.sNome
        
        For Each objCodigoNomeO In gcolFiliaisTela
            If objGeracaoNF.colNFFilialO.Item(iIndice) = objCodigoNomeO.iCodigo Then
                Exit For
            End If
        Next
    
        GridNF.TextMatrix(iIndice, iGrid_NFFilialEmpresa_Col) = objCodigoNomeO.iCodigo & SEPARADOR & objCodigoNomeO.sNome
        
        
        GridNF.TextMatrix(iIndice, iGrid_NFNumNota_Col) = CStr(objNF.lNumNotaFiscal)
        
        NFCliente.Text = CStr(objNF.lCliente)
        
        lErro = TP_Cliente_Le2(NFCliente, objcliente)
        If lErro <> SUCESSO Then gError 192256
        
        GridNF.TextMatrix(iIndice, iGrid_NFCliente_Col) = NFCliente.Text
        GridNF.TextMatrix(iIndice, iGrid_NFNumItens_Col) = CStr(objNF.ColItensNF.Count)
        GridNF.TextMatrix(iIndice, iGrid_NFValor_Col) = Format(objNF.dValorTotal, "STANDARD")
        
        For Each objItemNF In objNF.ColItensNF
        
            iItem = iItem + 1
        
            GridItens.TextMatrix(iItem, iGrid_ItemFilialEmpresa_Col) = objCodigoNomeD.iCodigo & SEPARADOR & objCodigoNomeD.sNome
            GridItens.TextMatrix(iItem, iGrid_ItemNumNota_Col) = CStr(objNF.lNumNotaFiscal)
            GridItens.TextMatrix(iItem, iGrid_Item_Col) = CStr(objItemNF.iItem)
            GridItens.TextMatrix(iItem, iGrid_ItemCliente_Col) = NFCliente.Text
            GridItens.TextMatrix(iItem, iGrid_ItemProduto_Col) = objItemNF.sProduto
            
            sParte = objItemNF.sDescricaoItem
    
            iPOS = InStr(1, sParte, " ")
    
            Do While iPOS <> 0
    
                sParte = Mid(sParte, iPOS + 1)
    
                iPOS = InStr(1, sParte, " ")
    
            Loop
            
            GridItens.TextMatrix(iItem, iGrid_ItemFat_Col) = sParte
            GridItens.TextMatrix(iItem, iGrid_ItemDesc_Col) = objItemNF.sDescricaoItem
            GridItens.TextMatrix(iItem, iGrid_ItemValor_Col) = Format(objItemNF.dValorTotal, "STANDARD")
        
        Next
    
    Next
    
    objGridNF.iLinhasExistentes = iIndice
    objGridItens.iLinhasExistentes = iItem
                
    GL_objMDIForm.MousePointer = vbDefault
                
    Traz_GeracaoNF_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_GeracaoNF_Tela:

    GL_objMDIForm.MousePointer = vbDefault

    Traz_GeracaoNF_Tela = gErr
    
    Select Case gErr

        Case 192253, 192254, 192256
              
        Case 192255
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_NENHUM_VOUCHER", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192257)

    End Select

End Function

Private Sub TipoDocSeleciona_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 192258

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 192258

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192259)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 192260

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 192260

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192261)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 192262

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 192262

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192263)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 192264

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 192264

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192265)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
  
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192266)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        If objGridInt.objGrid.Name = GridFE.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_FESel_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, FESel)
                    If lErro <> SUCESSO Then gError 196510
                    
                Case iGrid_FEPerc_Col
                
                    lErro = Saida_Celula_Percentual(objGridInt, FEPerc)
                    If lErro <> SUCESSO Then gError 196511
                    
            End Select
            
        End If
                    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 192267

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
    
        Case 196510, 196511

        Case 192267
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192268)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Filial(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("NFs")
    objGridInt.colColuna.Add ("Títulos")
    objGridInt.colColuna.Add ("Valor O.")
    objGridInt.colColuna.Add ("Valor R.")
    objGridInt.colColuna.Add ("% R.")
    objGridInt.colColuna.Add ("Notas: De")
    objGridInt.colColuna.Add ("Até")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    objGridInt.colCampo.Add (NumNF.Name)
    objGridInt.colCampo.Add (NumTitulos.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (ValorR.Name)
    objGridInt.colCampo.Add (PercR.Name)
    objGridInt.colCampo.Add (NFDe.Name)
    objGridInt.colCampo.Add (NFAte.Name)
    
    iGrid_FilialEmpresa_Col = 1
    iGrid_NumNF_Col = 2
    iGrid_NumTitulos_Col = 3
    iGrid_Valor_Col = 4
    iGrid_ValorR_Col = 5
    iGrid_PercR_Col = 6
    iGrid_NFDe_Col = 7
    iGrid_NFAte_Col = 8
    
    objGridInt.objGrid = GridFilial

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridFilial.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Private Function Inicializa_Grid_NF(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Filial O")
    objGridInt.colColuna.Add ("Filial D")
    objGridInt.colColuna.Add ("NF")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Itens")
    objGridInt.colColuna.Add ("Valor")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (NFFilialEmpresa.Name)
    objGridInt.colCampo.Add (NFFilialEmpresaR.Name)
    objGridInt.colCampo.Add (NFNumNota.Name)
    objGridInt.colCampo.Add (NFCliente.Name)
    objGridInt.colCampo.Add (NFNumItens.Name)
    objGridInt.colCampo.Add (NFValor.Name)

    iGrid_NFFilialEmpresa_Col = 1
    iGrid_NFFilialEmpresaR_Col = 2
    iGrid_NFNumNota_Col = 3
    iGrid_NFCliente_Col = 4
    iGrid_NFNumItens_Col = 5
    iGrid_NFValor_Col = 6
    
    objGridInt.objGrid = GridNF

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 15

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridNF.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("NF")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Prod.")
    objGridInt.colColuna.Add ("Fat.")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Valor")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (ItemFilialEmpresa.Name)
    objGridInt.colCampo.Add (ItemNumNota.Name)
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (ItemCliente.Name)
    objGridInt.colCampo.Add (ItemProduto.Name)
    objGridInt.colCampo.Add (ItemFat.Name)
    objGridInt.colCampo.Add (ItemDesc.Name)
    objGridInt.colCampo.Add (ItemValor.Name)

    iGrid_ItemFilialEmpresa_Col = 1
    iGrid_ItemNumNota_Col = 2
    iGrid_Item_Col = 3
    iGrid_ItemCliente_Col = 4
    iGrid_ItemProduto_Col = 5
    iGrid_ItemFat_Col = 6
    iGrid_ItemDesc_Col = 7
    iGrid_ItemValor_Col = 8
    
    objGridInt.objGrid = GridItens

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 15

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    If gobjGeracaoNF.colNF.Count = 0 Then gError 192269
    
    'Libera os Bloqueios selecionados
    lErro = CF("TRV_NF_Gera", gobjGeracaoNF)
    If lErro <> SUCESSO Then gError 192270
  
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 192269
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_NF_PARA_GERACAO", gErr)
        
        Case 192270

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192271)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria

    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192272)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Geração de Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TRVGeracaoNF"
    
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

Private Sub TabStripOpcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStripOpcao)
End Sub

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_FilialEmpresa

    'FiliaisEmpresa.Clear
    
    Set gcolFiliaisTela = New Collection

    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 192273

    iIndice = 0
    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
    
        gcolFiliaisTela.Add objCodigoNome
    
        If objCodigoNome.iCodigo < Abs(giFilialAuxiliar) Then
            
            'FiliaisEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            'FiliaisEmpresa.ItemData(FiliaisEmpresa.NewIndex) = objCodigoNome.iCodigo
            'FiliaisEmpresa.Selected(iIndice) = True
        
            iIndice = iIndice + 1
        
            GridFE.TextMatrix(iIndice, iGrid_FESel_Col) = CStr(MARCADO)
            GridFE.TextMatrix(iIndice, iGrid_FE_Col) = objCodigoNome.iCodigo & SEPARADOR & objCodigoNome.sNome
        
        End If
    
    Next
    
    objGridFE.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridFE)

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 192273

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192274)

    End Select

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    
    objGridInt.objGrid.Rows = iNumLinhas + 1
    
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
End Sub

Private Sub BotaoMarcarTodos_Click()

Dim iLinha As Integer

On Error GoTo Erro_BotaoMarcarTodos_Click
    
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    For iLinha = 1 To objGridFE.iLinhasExistentes
        GridFE.TextMatrix(iLinha, iGrid_FESel_Col) = CStr(MARCADO)
    Next
   
    Call Grid_Refresh_Checkbox(objGridFE)
    
    Exit Sub

Erro_BotaoMarcarTodos_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192275)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoDesmarcarTodos_Click()

Dim iLinha As Integer

On Error GoTo Erro_BotaoDesmarcarTodos_Click

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    For iLinha = 1 To objGridFE.iLinhasExistentes
        GridFE.TextMatrix(iLinha, iGrid_FESel_Col) = CStr(DESMARCADO)
    Next
   
    Call Grid_Refresh_Checkbox(objGridFE)

    Exit Sub

Erro_BotaoDesmarcarTodos_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192276)

    End Select
    
    Exit Sub
    
End Sub

Sub BotaoGerar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGerar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 192277

    'Limpa Tela
    Call Limpa_Tela_GeracaoNF

    Exit Sub

Erro_BotaoGerar_Click:

    Select Case gErr

        Case 192277

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192278)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 192279

    'Limpa o restante da tela
    Call Limpa_Tela_GeracaoNF
    
    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 192279
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192280)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_GeracaoNF()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_GeracaoNF

    Call Limpa_Tela(Me)

    'Limpa os Grids da tela
    Call Grid_Limpa(objGridFilial)
    Call Grid_Limpa(objGridNF)
    Call Grid_Limpa(objGridItens)
    'Call Grid_Limpa(objGridFE)
    
    Call Ordenacao_Limpa(objGridFilial)
    Call Ordenacao_Limpa(objGridNF)
    Call Ordenacao_Limpa(objGridItens)
    Call Ordenacao_Limpa(objGridFE)
    
    Set gobjGeracaoNF = Nothing
    Set gcolNFOrd = Nothing
    Set gcolTitOrd = Nothing

    ValorTotal.Caption = ""
    DescItem.Caption = ""
    
    Call Default_Tela

    Exit Sub

Erro_Limpa_Tela_GeracaoNF:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192281)

    End Select

    Exit Sub

End Sub

Sub Default_Tela()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Default_Tela
    
    OptEmitido.Value = True
    optExtraiDados.Value = vbUnchecked

    For iIndice = 1 To objGridFE.iLinhasExistentes
        GridFE.TextMatrix(iIndice, iGrid_FESel_Col) = CStr(MARCADO)
        GridFE.TextMatrix(iIndice, iGrid_FEPerc_Col) = ""
    Next
    
    Call Combo_Seleciona_ItemData(Marca, 0)
    
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Default_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192282)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 192303
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
                    
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 192303

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192304)

    End Select

    Exit Function

End Function

Public Sub TipoDocApenas_Click()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    'Habilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = True

End Sub

Public Sub TipoDocTodos_Click()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    'Desabilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = False

    'Limpa a combo de seleção de conta corrente
    TipoDocSeleciona.ListIndex = COMBO_INDICE

End Sub

Private Sub GridItens_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
    
        If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then

            'Seta objTela como a Tela de Baixas a Receber
            Set PopUpMenuGerNF.objTela = Me

            'Chama o Menu PopUp
            PopUpMenuGerNF.PopupMenu PopUpMenuGerNF.mnuGrid, vbPopupMenuRightButton
    
            'Limpa o objTela
            Set PopUpMenuGerNF.objTela = Nothing
            
        End If

    End If
    
End Sub

Public Function mnuTvwAbrirDoc_Click() As Long

Dim lErro As Long
Dim objTitRec As ClassTituloReceber

On Error GoTo Erro_mnuTvwAbrirDoc_Click

    Set objTitRec = gcolTitOrd(GridItens.Row)
    
    Call Chama_Tela(TRV_TIPO_DOC_DESTINO_TITREC_TELA, objTitRec)
    
    mnuTvwAbrirDoc_Click = SUCESSO
    
    Exit Function

Erro_mnuTvwAbrirDoc_Click:

    mnuTvwAbrirDoc_Click = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192376)

    End Select

    Exit Function
    
End Function

Private Sub FaturaDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FaturaDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(FaturaDe, iAlterado)

End Sub

Private Sub FaturaAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FaturaAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(FaturaAte, iAlterado)

End Sub

Private Sub UpDownDataBaixaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixaAte_DownClick

    'Diminui a DataBaixaAte em 1 dia
    lErro = Data_Up_Down_Click(DataBaixaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 196512

    Exit Sub

Erro_UpDownDataBaixaAte_DownClick:

    Select Case gErr

        Case 196512

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196513)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataBaixaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixaAte_UpClick

    'Aumenta a DataBaixaAte em 1 dia
    lErro = Data_Up_Down_Click(DataBaixaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 196514

    Exit Sub

Erro_UpDownDataBaixaAte_UpClick:

    Select Case gErr

        Case 196514

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196515)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataBaixaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixaDe_DownClick

    'Diminui a DataBaixaDe em 1 dia
    lErro = Data_Up_Down_Click(DataBaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 196516

    Exit Sub

Erro_UpDownDataBaixaDe_DownClick:

    Select Case gErr

        Case 196516

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196517)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataBaixaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixaDe_UpClick

    'Aumenta a DataBaixaDe em 1 dia
    lErro = Data_Up_Down_Click(DataBaixaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 196518

    Exit Sub

Erro_UpDownDataBaixaDe_UpClick:

    Select Case gErr

        Case 196518

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196519)

    End Select

    Exit Sub

End Sub

Private Sub DataBaixaAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataBaixaAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataBaixaAte, iAlterado)

End Sub

Private Sub DataBaixaAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBaixaAte_Validate

    'Se a DataBaixaAte está preenchida
    If Len(DataBaixaAte.ClipText) > 0 Then

        'Verifica se a DataBaixaAte é válida
        lErro = Data_Critica(DataBaixaAte.Text)
        If lErro <> SUCESSO Then gError 196520

    End If
    
    Exit Sub

Erro_DataBaixaAte_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 196520 'Tratado na rotina chamada

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196521)

    End Select

    Exit Sub

End Sub

Private Sub DataBaixaDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataBaixaDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataBaixaDe, iAlterado)

End Sub

Private Sub DataBaixaDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBaixaDe_Validate

    'Se a DataBaixaDe está preenchida
    If Len(DataBaixaDe.ClipText) > 0 Then

        'Verifica se a DataBaixaDe é válida
        lErro = Data_Critica(DataBaixaDe.Text)
        If lErro <> SUCESSO Then gError 196522

    End If

    Exit Sub

Erro_DataBaixaDe_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 196522

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196523)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_FE(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("%")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (FESel.Name)
    objGridInt.colCampo.Add (FE.Name)
    objGridInt.colCampo.Add (FEPerc.Name)

    iGrid_FESel_Col = 1
    iGrid_FE_Col = 2
    iGrid_FEPerc_Col = 3
    
    objGridInt.objGrid = GridFE

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridFE.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Public Sub FEPerc_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FEPerc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFE)
End Sub

Public Sub FEPerc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFE)
End Sub

Public Sub FEPerc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridFE.objControle = FEPerc
    lErro = Grid_Campo_Libera_Foco(objGridFE)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub FESel_Click()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FESel_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFE)
End Sub

Public Sub FESel_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFE)
End Sub

Public Sub FESel_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridFE.objControle = FESel
    lErro = Grid_Campo_Libera_Foco(objGridFE)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Function Saida_Celula_Percentual(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim dPercent As Double

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(objControle.Text)
        If lErro <> SUCESSO Then gError 196524

        dPercent = StrParaDbl(objControle.Text)

        'se for igual a 100% -> erro
        If dPercent = 100 Then gError 196525

        objControle.Text = Format(dPercent, "Fixed")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196526

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = gErr

    Select Case gErr

        Case 196524, 196526
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 196525
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196527)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196528

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 196528
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196529)

    End Select

    Exit Function

End Function

Private Sub GridFE_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFE, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFE, iAlterado)
    End If

End Sub

Private Sub GridFE_GotFocus()
    Call Grid_Recebe_Foco(objGridFE)
End Sub

Private Sub GridFE_EnterCell()
    Call Grid_Entrada_Celula(objGridFE, iAlterado)
End Sub

Private Sub GridFE_LeaveCell()
    Call Saida_Celula(objGridFE)
End Sub

Private Sub GridFE_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridFE)
End Sub

Private Sub GridFE_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFE, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFE, iAlterado)
    End If

End Sub

Private Sub GridFE_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridFE)
End Sub

Private Sub GridFE_RowColChange()
    Call Grid_RowColChange(objGridFE)
End Sub

Private Sub GridFE_Scroll()
    Call Grid_Scroll(objGridFE)
End Sub

Private Sub BotaoTitSemNF_Click(Index As Integer)


Dim sFiltro As String
Dim colSelecao As New Collection
Dim vValor As Variant

On Error GoTo Erro_BotaoTitSemNF

    sFiltro = gobjGeracaoNF.sFiltro
    
    For Each vValor In gobjGeracaoNF.colSelecao
        colSelecao.Add vValor
    Next

    Select Case Index

        Case TAB_GERACAO
        
            If GridFilial.Row = 0 Then gError 196533
        
            sFiltro = sFiltro & IIf(Len(Trim(sFiltro)) > 0, " AND ", "") & " FilialEmpresa = ?"
            colSelecao.Add Codigo_Extrai(GridFilial.TextMatrix(GridFilial.Row, iGrid_FilialEmpresa_Col))
        

        Case TAB_NF
            
            If GridNF.Row = 0 Then gError 196534
            
            sFiltro = sFiltro & IIf(Len(Trim(sFiltro)) > 0, " AND ", "") & " Cliente = ?"
            colSelecao.Add LCodigo_Extrai(GridNF.TextMatrix(GridNF.Row, iGrid_NFCliente_Col))

        Case Else


    End Select
    
    Call Chama_Tela("TitulosSemNotaLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoTitSemNF:
    
    Select Case gErr
    
        Case 196533, 196534
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196535)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoConsTitulo_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoTitSemNF_Click

    If GridItens.Row = 0 Then gError 196536

    Call mnuTvwAbrirDoc_Click

    Exit Sub
    
Erro_BotaoTitSemNF_Click:
    
    Select Case gErr
    
        Case 196536
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196537)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub UpDownDataEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_DownClick

    'Diminui a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 192262

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case gErr

        Case 192262

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192263)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_UpClick

    'Aumenta a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 192264

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case gErr

        Case 192264

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192265)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Se a DataEmissao está preenchida
    If Len(DataEmissao.ClipText) > 0 Then

        'Verifica se a DataEmissao é válida
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError 192242

    End If

    Exit Sub

Erro_DataEmissao_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 192242

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192243)

    End Select

    Exit Sub

End Sub

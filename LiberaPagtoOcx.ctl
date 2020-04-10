VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl LiberaPagtoOcx 
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12255
   ScaleHeight     =   7500
   ScaleWidth      =   12255
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5625
      Index           =   2
      Left            =   480
      TabIndex        =   32
      Top             =   1365
      Visible         =   0   'False
      Width           =   11385
      Begin VB.ComboBox Ordenacao 
         Height          =   315
         ItemData        =   "LiberaPagtoOcx.ctx":0000
         Left            =   1770
         List            =   "LiberaPagtoOcx.ctx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   90
         Width           =   3015
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   600
         Left            =   4650
         Picture         =   "LiberaPagtoOcx.ctx":0087
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4485
         Width           =   1440
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   600
         Left            =   2835
         Picture         =   "LiberaPagtoOcx.ctx":1269
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4485
         Width           =   1440
      End
      Begin VB.CommandButton BotaoGravar 
         Caption         =   "Libera os Pagamentos Assinalados"
         Height          =   1185
         Left            =   660
         Picture         =   "LiberaPagtoOcx.ctx":2283
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4380
         Width           =   1590
      End
      Begin VB.CommandButton BotaoConsultaDocOriginal 
         Height          =   450
         Left            =   8895
         Picture         =   "LiberaPagtoOcx.ctx":26C5
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Consulta o documento original de uma parcela"
         Top             =   4380
         Width           =   1065
      End
      Begin VB.Frame FrameParcelas 
         Caption         =   "Parcelas em Aberto"
         Height          =   3690
         Left            =   60
         TabIndex        =   33
         Top             =   465
         Width           =   11220
         Begin VB.CheckBox Selecionada 
            Height          =   220
            Left            =   8025
            TabIndex        =   34
            Top             =   270
            Width           =   570
         End
         Begin MSMask.MaskEdBox DataEmissaoTitulo 
            Height          =   225
            Left            =   4530
            TabIndex        =   35
            Top             =   1080
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
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   4485
            TabIndex        =   36
            Top             =   570
            Width           =   1290
            _ExtentX        =   2275
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
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   360
            TabIndex        =   37
            Top             =   225
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
         Begin MSMask.MaskEdBox Saldo 
            Height          =   225
            Left            =   3000
            TabIndex        =   38
            Top             =   255
            Width           =   975
            _ExtentX        =   1720
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
         Begin MSMask.MaskEdBox Numero 
            Height          =   225
            Left            =   1785
            TabIndex        =   39
            Top             =   240
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Tipo 
            Height          =   225
            Left            =   1275
            TabIndex        =   40
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Parcela 
            Height          =   225
            Left            =   2640
            TabIndex        =   41
            Top             =   255
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "99"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2430
            Left            =   120
            TabIndex        =   42
            Top             =   225
            Width           =   11010
            _ExtentX        =   19420
            _ExtentY        =   4286
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox FilialFornItem 
            Height          =   225
            Left            =   1530
            TabIndex        =   43
            Top             =   135
            Width           =   2025
            _ExtentX        =   3572
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
         Begin MSMask.MaskEdBox FornItem 
            Height          =   225
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   2025
            _ExtentX        =   3572
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
      End
      Begin MSMask.MaskEdBox NomePortador 
         Height          =   225
         Left            =   6510
         TabIndex        =   46
         Top             =   615
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cobranca 
         Height          =   225
         Left            =   4020
         TabIndex        =   47
         Top             =   1425
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FilialEmpresa 
         Height          =   255
         Left            =   7560
         TabIndex        =   48
         Top             =   600
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   "9999"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
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
         Left            =   6495
         TabIndex        =   55
         Top             =   4425
         Width           =   510
      End
      Begin VB.Label TotalBaixar 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7080
         TabIndex        =   54
         Top             =   4395
         Width           =   1560
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
         Left            =   390
         TabIndex        =   49
         Top             =   135
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   1
      Left            =   1425
      TabIndex        =   0
      Top             =   1560
      Width           =   9120
      Begin VB.Frame Frame8 
         Caption         =   "Fornecedor"
         Height          =   1005
         Left            =   255
         TabIndex        =   25
         Top             =   375
         Width           =   8355
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5475
            TabIndex        =   26
            Top             =   390
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1560
            TabIndex        =   27
            Top             =   397
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label Label12 
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
            Left            =   4920
            TabIndex        =   29
            Top             =   450
            Width           =   465
         End
         Begin VB.Label FornecLabel 
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
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   450
            Width           =   1035
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Filtros"
         Height          =   3585
         Left            =   255
         TabIndex        =   1
         Top             =   1455
         Width           =   8355
         Begin VB.Frame Frame4 
            Caption         =   "Data de Emissão"
            Height          =   1575
            Left            =   390
            TabIndex        =   18
            Top             =   270
            Width           =   2175
            Begin MSComCtl2.UpDown UpDownEmissaoInic 
               Height          =   300
               Left            =   1725
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   450
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoInic 
               Height          =   300
               Left            =   660
               TabIndex        =   20
               Top             =   465
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoFim 
               Height          =   300
               Left            =   1725
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   960
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoFim 
               Height          =   300
               Left            =   645
               TabIndex        =   22
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
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
               Left            =   240
               TabIndex        =   24
               Top             =   495
               Width           =   315
            End
            Begin VB.Label Label16 
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
               Left            =   195
               TabIndex        =   23
               Top             =   1013
               Width           =   360
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data de Vencimento"
            Height          =   1575
            Left            =   3150
            TabIndex        =   11
            Top             =   270
            Width           =   2175
            Begin MSComCtl2.UpDown UpDownVencInic 
               Height          =   300
               Left            =   1695
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   480
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencInic 
               Height          =   300
               Left            =   630
               TabIndex        =   13
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownVencFim 
               Height          =   300
               Left            =   1695
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   990
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencFim 
               Height          =   300
               Left            =   615
               TabIndex        =   15
               Top             =   990
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label17 
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
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   510
               Width           =   375
            End
            Begin VB.Label Label20 
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
               Height          =   255
               Left            =   210
               TabIndex        =   16
               Top             =   1020
               Width           =   375
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Nº do Título"
            Height          =   1575
            Left            =   5790
            TabIndex        =   6
            Top             =   270
            Width           =   2175
            Begin MSMask.MaskEdBox TituloInic 
               Height          =   300
               Left            =   720
               TabIndex        =   7
               Top             =   435
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TituloFim 
               Height          =   300
               Left            =   735
               TabIndex        =   8
               Top             =   960
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin VB.Label Label21 
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
               Height          =   255
               Left            =   360
               TabIndex        =   10
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label22 
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
               Height          =   255
               Left            =   315
               TabIndex        =   9
               Top             =   1005
               Width           =   375
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de Documento"
            Height          =   1410
            Left            =   390
            TabIndex        =   2
            Top             =   1950
            Width           =   4935
            Begin VB.ComboBox TipoDocSeleciona 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "LiberaPagtoOcx.ctx":35CF
               Left            =   1140
               List            =   "LiberaPagtoOcx.ctx":35D1
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   930
               Width           =   3510
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
               Left            =   75
               TabIndex        =   4
               Top             =   360
               Value           =   -1  'True
               Width           =   1005
            End
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
               Left            =   90
               TabIndex        =   3
               Top             =   960
               Width           =   1050
            End
         End
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
      Left            =   9765
      Picture         =   "LiberaPagtoOcx.ctx":35D3
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Fechar"
      Top             =   90
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   6705
      Left            =   240
      TabIndex        =   30
      Top             =   645
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   11827
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Liberação"
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
Attribute VB_Name = "LiberaPagtoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTLiberaPagto
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoConsultaDocOriginal_Click()
    Call objCT.BotaoConsultaDocOriginal_Click
End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call objCT.BotaoDesmarcarTodos_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoMarcarTodos_Click()
    Call objCT.BotaoMarcarTodos_Click
End Sub

Private Sub Cobranca_GotFocus()
     Call objCT.Cobranca_GotFocus
End Sub

Private Sub Cobranca_KeyPress(KeyAscii As Integer)
     Call objCT.Cobranca_KeyPress(KeyAscii)
End Sub

Private Sub Cobranca_Validate(Cancel As Boolean)
     Call objCT.Cobranca_Validate(Cancel)
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimento_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
End Sub

Private Sub EmissaoFim_Change()
     Call objCT.EmissaoFim_Change
End Sub

Private Sub EmissaoFim_GotFocus()
     Call objCT.EmissaoFim_GotFocus
End Sub

Private Sub EmissaoFim_Validate(Cancel As Boolean)
     Call objCT.EmissaoFim_Validate(Cancel)
End Sub

Private Sub EmissaoInic_Change()
     Call objCT.EmissaoInic_Change
End Sub

Private Sub EmissaoInic_GotFocus()
     Call objCT.EmissaoInic_GotFocus
End Sub

Private Sub EmissaoInic_Validate(Cancel As Boolean)
     Call objCT.EmissaoInic_Validate(Cancel)
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

Private Sub FilialEmpresa_GotFocus()
     Call objCT.FilialEmpresa_GotFocus
End Sub

Private Sub FilialEmpresa_KeyPress(KeyAscii As Integer)
     Call objCT.FilialEmpresa_KeyPress(KeyAscii)
End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)
     Call objCT.FilialEmpresa_Validate(Cancel)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Fornecedor_Change()
     Call objCT.Fornecedor_Change
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
     Call objCT.Fornecedor_Validate(Cancel)
End Sub

Private Sub FornecLabel_Click()
     Call objCT.FornecLabel_Click
End Sub

Public Sub mnuGridConsultaDocOriginal_Click()
    Call objCT.mnuGridConsultaDocOriginal_Click
End Sub

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call objCT.GridParcelas_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub Numero_GotFocus()
     Call objCT.Numero_GotFocus
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
     Call objCT.Numero_KeyPress(KeyAscii)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
     Call objCT.Numero_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Parcela_GotFocus()
     Call objCT.Parcela_GotFocus
End Sub

Private Sub Parcela_KeyPress(KeyAscii As Integer)
     Call objCT.Parcela_KeyPress(KeyAscii)
End Sub

Private Sub Parcela_Validate(Cancel As Boolean)
     Call objCT.Parcela_Validate(Cancel)
End Sub

Private Sub Saldo_GotFocus()
     Call objCT.Saldo_GotFocus
End Sub

Private Sub Saldo_KeyPress(KeyAscii As Integer)
     Call objCT.Saldo_KeyPress(KeyAscii)
End Sub

Private Sub Saldo_Validate(Cancel As Boolean)
     Call objCT.Saldo_Validate(Cancel)
End Sub

Private Sub Selecionada_Click()
     Call objCT.Selecionada_Click
End Sub

Private Sub Tipo_GotFocus()
     Call objCT.Tipo_GotFocus
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)
     Call objCT.Tipo_KeyPress(KeyAscii)
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub TituloFim_Change()
     Call objCT.TituloFim_Change
End Sub

Private Sub TituloFim_GotFocus()
     Call objCT.TituloFim_GotFocus
End Sub

Private Sub TituloFim_Validate(Cancel As Boolean)
     Call objCT.TituloFim_Validate(Cancel)
End Sub

Private Sub TituloInic_Change()
     Call objCT.TituloInic_Change
End Sub

Private Sub TituloInic_GotFocus()
     Call objCT.TituloInic_GotFocus
End Sub

Private Sub UpDownEmissaoFim_DownClick()
     Call objCT.UpDownEmissaoFim_DownClick
End Sub

Private Sub UpDownEmissaoFim_UpClick()
     Call objCT.UpDownEmissaoFim_UpClick
End Sub

Private Sub UpDownEmissaoInic_DownClick()
     Call objCT.UpDownEmissaoInic_DownClick
End Sub

Private Sub UpDownEmissaoInic_UpClick()
     Call objCT.UpDownEmissaoInic_UpClick
End Sub

Private Sub UpDownVencFim_DownClick()
     Call objCT.UpDownVencFim_DownClick
End Sub

Private Sub UpDownVencFim_UpClick()
     Call objCT.UpDownVencFim_UpClick
End Sub

Private Sub UpDownVencInic_DownClick()
     Call objCT.UpDownVencInic_DownClick
End Sub

Private Sub UpDownVencInic_UpClick()
     Call objCT.UpDownVencInic_UpClick
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTLiberaPagto
    Set objCT.objUserControl = Me
End Sub

Private Sub Selecionada_GotFocus()
     Call objCT.Selecionada_GotFocus
End Sub

Private Sub Selecionada_KeyPress(KeyAscii As Integer)
     Call objCT.Selecionada_KeyPress(KeyAscii)
End Sub

Private Sub Selecionada_Validate(Cancel As Boolean)
     Call objCT.Selecionada_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub VencFim_Change()
     Call objCT.VencFim_Change
End Sub

Private Sub VencFim_GotFocus()
     Call objCT.VencFim_GotFocus
End Sub

Private Sub VencFim_Validate(Cancel As Boolean)
     Call objCT.VencFim_Validate(Cancel)
End Sub

Private Sub VencInic_Change()
     Call objCT.VencInic_Change
End Sub

Private Sub VencInic_GotFocus()
     Call objCT.VencInic_GotFocus
End Sub

Private Sub VencInic_Validate(Cancel As Boolean)
     Call objCT.VencInic_Validate(Cancel)
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_UnLoad(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
        If Cancel = False Then
            Set objCT.objUserControl = Nothing
            Set objCT = Nothing
        End If
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



Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub FornecLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecLabel, Source, X, Y)
End Sub

Private Sub FornecLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub TipoDocTodos_Click()
     Call objCT.TipoDocTodos_Click
End Sub

Private Sub TipoDocApenas_Click()
     Call objCT.TipoDocApenas_Click
End Sub

Private Sub TipoDocSeleciona_Change()
     Call objCT.TipoDocSeleciona_Change
End Sub

Private Sub TipoDocSeleciona_Click()
     Call objCT.TipoDocSeleciona_Change
End Sub

Private Sub Ordenacao_Change()
    objCT.Ordenacao_Change
End Sub

Private Sub Ordenacao_Click()
    objCT.Ordenacao_Click
End Sub


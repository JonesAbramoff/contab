VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaCartaoOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   2
      Left            =   60
      TabIndex        =   47
      Top             =   645
      Visible         =   0   'False
      Width           =   9405
      Begin VB.Frame Frame2 
         Caption         =   "Arquivos"
         Height          =   5175
         Index           =   2
         Left            =   0
         TabIndex        =   48
         Top             =   45
         Width           =   9405
         Begin VB.CheckBox ArqIgnorar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   215
            Left            =   405
            TabIndex        =   112
            ToolTipText     =   "Marca o arquivo para ser ignorado"
            Top             =   1125
            Width           =   435
         End
         Begin VB.CheckBox ArqSelecionado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   215
            Left            =   420
            TabIndex        =   73
            ToolTipText     =   "Seleciona o arquivo para importação"
            Top             =   885
            Width           =   435
         End
         Begin MSMask.MaskEdBox ArqVlrDep 
            Height          =   210
            Left            =   6240
            TabIndex        =   68
            ToolTipText     =   "Valor dos depósitos informados no arquivo"
            Top             =   870
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox ArqHoraImp 
            Height          =   210
            Left            =   3300
            TabIndex        =   49
            ToolTipText     =   "Hora da Importação"
            Top             =   870
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox ArqDataImp 
            Height          =   210
            Left            =   2445
            TabIndex        =   50
            ToolTipText     =   "Data da Importação"
            Top             =   870
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox ArqNomeArq 
            Height          =   210
            Left            =   855
            TabIndex        =   51
            Top             =   870
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox ArqQtdDep 
            Height          =   210
            Left            =   4080
            TabIndex        =   52
            ToolTipText     =   "Quantidades de depósitos dentro do arquivo"
            Top             =   870
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   370
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
            Format          =   "#,##0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ArqQtdParc 
            Height          =   210
            Left            =   4800
            TabIndex        =   54
            ToolTipText     =   "Quantidades de parcelas dentro do arquivo"
            Top             =   870
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   370
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
            Format          =   "#,##0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ArqQtdParcEnc 
            Height          =   210
            Left            =   5520
            TabIndex        =   55
            ToolTipText     =   "Quantidades de parcelas localizadas no sistema"
            Top             =   870
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   370
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
            Format          =   "#,##0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ArqVlrParc 
            Height          =   210
            Left            =   7125
            TabIndex        =   56
            ToolTipText     =   "Valor das parcelas informadas no arquivo"
            Top             =   870
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox ArqVlrParcEnc 
            Height          =   210
            Left            =   8010
            TabIndex        =   57
            ToolTipText     =   "Valor das parcelas localizadas no sistema"
            Top             =   870
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   370
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
         Begin VB.Frame Frame2 
            Caption         =   "Resumo"
            Height          =   1080
            Index           =   8
            Left            =   90
            TabIndex        =   53
            Top             =   4035
            Width           =   9225
            Begin VB.Frame Frame2 
               Caption         =   "Sistema"
               Height          =   855
               Index           =   9
               Left            =   6135
               TabIndex        =   63
               Top             =   165
               Width           =   3015
               Begin VB.Label ArqTotalQtdParcEnc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   66
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label ArqTotalVlrParcEnc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   64
                  Top             =   510
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Parcelas:"
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
                  Index           =   8
                  Left            =   165
                  TabIndex        =   67
                  Top             =   225
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Parcelas:"
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
                  Index           =   7
                  Left            =   135
                  TabIndex        =   65
                  Top             =   540
                  Width           =   1515
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Arquivos"
               Height          =   855
               Index           =   7
               Left            =   120
               TabIndex        =   58
               Top             =   165
               Width           =   5985
               Begin VB.Label ArqTotalVlrDep 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   4575
                  TabIndex        =   72
                  Top             =   510
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Depósitos:"
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
                  Index           =   10
                  Left            =   3015
                  TabIndex        =   71
                  Top             =   540
                  Width           =   1515
               End
               Begin VB.Label ArqTotalQtdDep 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   4575
                  TabIndex        =   70
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Depósitos:"
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
                  Index           =   9
                  Left            =   3045
                  TabIndex        =   69
                  Top             =   225
                  Width           =   1515
               End
               Begin VB.Label ArqTotalVlrParc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   62
                  Top             =   510
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Parcelas:"
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
                  Index           =   6
                  Left            =   135
                  TabIndex        =   61
                  Top             =   540
                  Width           =   1515
               End
               Begin VB.Label ArqTotalQtdParc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   60
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Parcelas:"
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
                  Index           =   4
                  Left            =   165
                  TabIndex        =   59
                  Top             =   225
                  Width           =   1515
               End
            End
         End
         Begin MSFlexGridLib.MSFlexGrid GridArq 
            Height          =   390
            Left            =   30
            TabIndex        =   9
            Top             =   240
            Width           =   9345
            _ExtentX        =   16484
            _ExtentY        =   688
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5205
      Index           =   1
      Left            =   60
      TabIndex        =   20
      Top             =   660
      Width           =   9390
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   3675
         Index           =   5
         Left            =   1215
         TabIndex        =   33
         Top             =   345
         Width           =   6765
         Begin VB.Frame Frame2 
            Caption         =   "Bandeira"
            Height          =   1320
            Index           =   3
            Left            =   495
            TabIndex        =   34
            Top             =   570
            Width           =   5700
            Begin VB.CheckBox Cielo 
               Caption         =   "Arquivo da Cielo"
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
               Left            =   1200
               TabIndex        =   1
               Top             =   795
               Width           =   2550
            End
            Begin VB.ComboBox Bandeira 
               Height          =   315
               ItemData        =   "BaixasCartao.ctx":0000
               Left            =   1200
               List            =   "BaixasCartao.ctx":000D
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   345
               Width           =   3720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Bandeira:"
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
               Left            =   300
               TabIndex        =   35
               Top             =   375
               Width           =   825
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Período de importação"
            Height          =   885
            Index           =   6
            Left            =   495
            TabIndex        =   36
            Top             =   2055
            Width           =   5715
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   315
               Left            =   2325
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   375
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   315
               Left            =   1170
               TabIndex        =   2
               Top             =   375
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   315
               Left            =   4650
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   375
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   315
               Left            =   3510
               TabIndex        =   4
               Top             =   375
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
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
               Index           =   3
               Left            =   3105
               TabIndex        =   38
               Top             =   435
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   780
               TabIndex        =   37
               Top             =   435
               Width           =   315
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5205
      Index           =   4
      Left            =   60
      TabIndex        =   21
      Top             =   660
      Visible         =   0   'False
      Width           =   9420
      Begin VB.CommandButton BotaoExibeParcCartao 
         Caption         =   "Exibe todas as parcelas de cartão em aberto"
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
         Left            =   1500
         TabIndex        =   16
         Top             =   4740
         Width           =   2610
      End
      Begin VB.CommandButton BotaoExibeParcDataValor 
         Caption         =   "Exibe as compras feitas com data e valor próximos"
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
         Left            =   4140
         TabIndex        =   17
         Top             =   4740
         Width           =   2610
      End
      Begin VB.CommandButton BotaoExibeParcEsseCartao 
         Caption         =   "Exibe as compras feitas com esse cartão ou autorização"
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
         Left            =   6780
         TabIndex        =   18
         Top             =   4740
         Width           =   2610
      End
      Begin VB.CommandButton BotaoTitulo 
         Caption         =   "Consulta o Título"
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
         Left            =   15
         TabIndex        =   15
         Top             =   4740
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Lista de Parcelas"
         Height          =   4665
         Index           =   1
         Left            =   30
         TabIndex        =   23
         Top             =   60
         Width           =   9375
         Begin VB.Frame Frame3 
            Caption         =   "Detalhamento do Status"
            Height          =   540
            Left            =   75
            TabIndex        =   109
            Top             =   3000
            Width           =   9240
            Begin VB.TextBox Status 
               Height          =   285
               Left            =   105
               Locked          =   -1  'True
               TabIndex        =   110
               Top             =   195
               Width           =   9030
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Resumo"
            Height          =   1080
            Index           =   13
            Left            =   75
            TabIndex        =   93
            Top             =   3540
            Width           =   9240
            Begin VB.Frame Frame2 
               Caption         =   "Diferença"
               Height          =   855
               Index           =   16
               Left            =   6150
               TabIndex        =   104
               Top             =   195
               Width           =   3015
               Begin VB.Label DetTotalVlrParcDif 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   106
                  Top             =   510
                  Width           =   1350
               End
               Begin VB.Label DetTotalQtdParcDif 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   105
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Parcelas:"
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
                  Index           =   23
                  Left            =   135
                  TabIndex        =   108
                  Top             =   540
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Parcelas:"
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
                  Index           =   22
                  Left            =   165
                  TabIndex        =   107
                  Top             =   225
                  Width           =   1515
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Sistema"
               Height          =   855
               Index           =   15
               Left            =   3105
               TabIndex        =   99
               Top             =   180
               Width           =   3015
               Begin VB.Label DetTotalVlrParcEnc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   101
                  Top             =   510
                  Width           =   1350
               End
               Begin VB.Label DetTotalQtdParcEnc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   100
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Parcelas:"
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
                  Index           =   21
                  Left            =   135
                  TabIndex        =   103
                  Top             =   540
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Parcelas:"
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
                  Index           =   20
                  Left            =   165
                  TabIndex        =   102
                  Top             =   225
                  Width           =   1515
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Arquivo"
               Height          =   855
               Index           =   14
               Left            =   60
               TabIndex        =   94
               Top             =   180
               Width           =   3015
               Begin VB.Label DetTotalVlrParc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   96
                  Top             =   510
                  Width           =   1350
               End
               Begin VB.Label DetTotalQtdParc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   95
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Parcelas:"
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
                  Index           =   19
                  Left            =   135
                  TabIndex        =   98
                  Top             =   540
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Parcelas:"
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
                  Index           =   18
                  Left            =   165
                  TabIndex        =   97
                  Top             =   225
                  Width           =   1515
               End
            End
         End
         Begin VB.ComboBox Deposito 
            Height          =   315
            ItemData        =   "BaixasCartao.ctx":0036
            Left            =   3495
            List            =   "BaixasCartao.ctx":0038
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   210
            Width           =   5715
         End
         Begin VB.ComboBox Exibir 
            Height          =   315
            ItemData        =   "BaixasCartao.ctx":003A
            Left            =   735
            List            =   "BaixasCartao.ctx":0047
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   210
            Width           =   1740
         End
         Begin MSMask.MaskEdBox DetNumCartao 
            Height          =   210
            Left            =   1185
            TabIndex        =   24
            Top             =   1500
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   370
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
         Begin MSFlexGridLib.MSFlexGrid GridDet 
            Height          =   300
            Left            =   30
            TabIndex        =   14
            Top             =   570
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   529
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox DetDataCompra 
            Height          =   210
            Left            =   165
            TabIndex        =   41
            Top             =   1500
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox DetNumAuto 
            Height          =   210
            Left            =   3030
            TabIndex        =   42
            Top             =   1500
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox DetParcela 
            Height          =   210
            Left            =   4035
            TabIndex        =   43
            Top             =   1500
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   370
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
            Format          =   "#,##0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DetValor 
            Height          =   210
            Left            =   4935
            TabIndex        =   44
            Top             =   1500
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox DetStatus 
            Height          =   210
            Left            =   5925
            TabIndex        =   45
            Top             =   1500
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   370
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
            Format          =   "#,##0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DetNumTitulo 
            Height          =   210
            Left            =   6660
            TabIndex        =   46
            Top             =   1500
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   370
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
            Format          =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DetVlrParc 
            Height          =   210
            Left            =   7665
            TabIndex        =   111
            Top             =   1500
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   370
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
         Begin VB.Label Label1 
            Caption         =   "Depósito:"
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
            Left            =   2655
            TabIndex        =   40
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Exibir:"
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
            Index           =   1
            Left            =   165
            TabIndex        =   39
            Top             =   255
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5250
      Index           =   3
      Left            =   75
      TabIndex        =   25
      Top             =   645
      Visible         =   0   'False
      Width           =   9390
      Begin VB.Frame Frame2 
         Caption         =   "Depósitos"
         Height          =   5160
         Index           =   4
         Left            =   0
         TabIndex        =   26
         Top             =   45
         Width           =   9390
         Begin VB.Frame Frame2 
            Caption         =   "Resumo"
            Height          =   1080
            Index           =   10
            Left            =   90
            TabIndex        =   78
            Top             =   4035
            Width           =   9225
            Begin VB.Frame Frame2 
               Caption         =   "Arquivos"
               Height          =   855
               Index           =   12
               Left            =   120
               TabIndex        =   84
               Top             =   165
               Width           =   5985
               Begin VB.Label DepTotalQtdParc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   91
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label DepTotalVlrParc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   89
                  Top             =   510
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Parcelas:"
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
                  Index           =   17
                  Left            =   165
                  TabIndex        =   92
                  Top             =   225
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Parcelas:"
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
                  Index           =   16
                  Left            =   135
                  TabIndex        =   90
                  Top             =   540
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Depósitos:"
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
                  Index           =   15
                  Left            =   3045
                  TabIndex        =   88
                  Top             =   225
                  Width           =   1515
               End
               Begin VB.Label DepTotalQtdDep 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   4575
                  TabIndex        =   87
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Depósitos:"
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
                  Index           =   14
                  Left            =   3015
                  TabIndex        =   86
                  Top             =   540
                  Width           =   1515
               End
               Begin VB.Label DepTotalVlrDep 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   4575
                  TabIndex        =   85
                  Top             =   510
                  Width           =   1350
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Sistema"
               Height          =   855
               Index           =   11
               Left            =   6135
               TabIndex        =   79
               Top             =   165
               Width           =   3015
               Begin VB.Label DepTotalVlrParcEnc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   81
                  Top             =   510
                  Width           =   1350
               End
               Begin VB.Label DepTotalQtdParcEnc 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1590
                  TabIndex        =   80
                  Top             =   195
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor Parcelas:"
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
                  Index           =   13
                  Left            =   135
                  TabIndex        =   83
                  Top             =   540
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Caption         =   "Qtde Parcelas:"
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
                  Index           =   12
                  Left            =   165
                  TabIndex        =   82
                  Top             =   225
                  Width           =   1515
               End
            End
         End
         Begin VB.ComboBox Arquivo 
            Height          =   315
            ItemData        =   "BaixasCartao.ctx":0068
            Left            =   990
            List            =   "BaixasCartao.ctx":006A
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   225
            Width           =   3510
         End
         Begin MSMask.MaskEdBox DepDataDep 
            Height          =   210
            Left            =   2850
            TabIndex        =   27
            Top             =   1110
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox DepQtdParc 
            Height          =   210
            Left            =   7605
            TabIndex        =   28
            Top             =   1110
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   370
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
            Format          =   "#,##0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DepCtaCorr 
            Height          =   210
            Left            =   1545
            TabIndex        =   29
            Top             =   1425
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox DepEstab 
            Height          =   210
            Left            =   390
            TabIndex        =   30
            Top             =   1110
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   370
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
         Begin MSFlexGridLib.MSFlexGrid GridDep 
            Height          =   390
            Left            =   30
            TabIndex        =   11
            Top             =   600
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   688
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox DepVlrBruto 
            Height          =   210
            Left            =   3765
            TabIndex        =   31
            Top             =   1110
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox DepQtdParcEnc 
            Height          =   210
            Left            =   8325
            TabIndex        =   32
            Top             =   1110
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   370
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
            Format          =   "#,##0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DepVlrComis 
            Height          =   210
            Left            =   4725
            TabIndex        =   75
            Top             =   1110
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox DepVlrLiq 
            Height          =   210
            Left            =   5685
            TabIndex        =   76
            Top             =   1110
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   370
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
         Begin MSMask.MaskEdBox DepVlrParcEnc 
            Height          =   210
            Left            =   6645
            TabIndex        =   77
            Top             =   1110
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   370
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
         Begin VB.Label Label1 
            Caption         =   "Arquivo:"
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
            Index           =   11
            Left            =   225
            TabIndex        =   74
            Top             =   270
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   7830
      ScaleHeight     =   480
      ScaleWidth      =   1605
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   30
      Width           =   1665
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "BaixasCartao.ctx":006C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "BaixasCartao.ctx":01C6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "BaixasCartao.ctx":06F8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5595
      Left            =   15
      TabIndex        =   19
      Top             =   330
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Arquivos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Depósitos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Detalhamentos"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "BaixaCartaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Const TAB_TELA_SELECAO = 1
Private Const TAB_TELA_ARQUIVO = 2
Private Const TAB_TELA_DEPOSITO = 3
Private Const TAB_TELA_DETALHE = 4

Private Const TELA_EXIBE_TODOS = 0
Private Const TELA_COM_ERROS = 1
Private Const TELA_SEM_ERROS = 2

'Variáveis Globais
Dim iFrameAtual As Integer

Dim iAlterado As Integer

Dim gobjBaixaCartao As New ClassBaixaCartao
Dim gobjArq As New ClassAdmExtFinArqsLidos
Dim gobjMov As New ClassAdmExtFinMov
Dim colDet As New Collection

Dim gbLimpandoTela As Boolean

'GridArq
Dim objGridArq As AdmGrid
Dim iGrid_ArqSelecionado_Col As Integer
Dim iGrid_ArqIgnorar_Col As Integer
Dim iGrid_ArqNomeArq_Col As Integer
Dim iGrid_ArqDataImp_Col As Integer
Dim iGrid_ArqHoraImp_Col As Integer
Dim iGrid_ArqQtdDep_Col As Integer
Dim iGrid_ArqQtdParc_Col As Integer
Dim iGrid_ArqQtdParcEnc_Col As Integer
Dim iGrid_ArqVlrDep_Col As Integer
Dim iGrid_ArqVlrParc_Col As Integer
Dim iGrid_ArqVlrParcEnc_Col As Integer

'GridDep
Dim objGridDep As AdmGrid
Dim iGrid_DepEstab_Col As Integer
Dim iGrid_DepCtaCorr_Col As Integer
Dim iGrid_DepDataDep_Col As Integer
Dim iGrid_DepVlrBruto_Col As Integer
Dim iGrid_DepVlrComis_Col As Integer
Dim iGrid_DepVlrLiq_Col As Integer
Dim iGrid_DepVlrParcEnc_Col As Integer
Dim iGrid_DepQtdParc_Col As Integer
Dim iGrid_DepQtdParcEnc_Col As Integer

'GridDet
Dim objGridDet As AdmGrid
Dim iGrid_DetDataCompra_Col As Integer
Dim iGrid_DetNumCartao_Col As Integer
Dim iGrid_DetNumAuto_Col As Integer
Dim iGrid_DetParcela_Col As Integer
Dim iGrid_DetValor_Col As Integer
Dim iGrid_DetStatus_Col As Integer
Dim iGrid_DetNumTitulo_Col As Integer
Dim iGrid_DetVlrParc_Col As Integer

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = TAB_TELA_SELECAO
    
    Set objGridArq = New AdmGrid
    Set objGridDep = New AdmGrid
    Set objGridDet = New AdmGrid
    
    lErro = Inicializa_Grid_Arq(objGridArq)
    If lErro <> SUCESSO Then gError 202605

    lErro = Inicializa_Grid_Dep(objGridDep)
    If lErro <> SUCESSO Then gError 202606

    lErro = Inicializa_Grid_Det(objGridDet)
    If lErro <> SUCESSO Then gError 202607

    Call Default_Tela
    
    gbLimpandoTela = False
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr
    
        Case 202605 To 202607

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202608)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set objGridArq = Nothing
    Set objGridDep = Nothing
    Set objGridDet = Nothing

    Set gobjBaixaCartao = Nothing
    Set gobjArq = Nothing
    Set gobjMov = Nothing
    Set colDet = Nothing

End Sub

Private Function Inicializa_Grid_Arq(objGridInt As AdmGrid) As Long
'Executa a Inicialização do gridarq

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Arq

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("S")
    objGridInt.colColuna.Add ("I")
    objGridInt.colColuna.Add ("Nome Arquivo")
    objGridInt.colColuna.Add ("Data Imp.")
    objGridInt.colColuna.Add ("Hora Imp.")
    objGridInt.colColuna.Add ("Depó.")
    objGridInt.colColuna.Add ("Parcs")
    objGridInt.colColuna.Add ("Sist.")
    objGridInt.colColuna.Add ("Depósit.R$")
    objGridInt.colColuna.Add ("Parcela.R$")
    objGridInt.colColuna.Add ("Sistema R$")

    'campos de edição do grid
    objGridInt.colCampo.Add (ArqSelecionado.Name)
    objGridInt.colCampo.Add (ArqIgnorar.Name)
    objGridInt.colCampo.Add (ArqNomeArq.Name)
    objGridInt.colCampo.Add (ArqDataImp.Name)
    objGridInt.colCampo.Add (ArqHoraImp.Name)
    objGridInt.colCampo.Add (ArqQtdDep.Name)
    objGridInt.colCampo.Add (ArqQtdParc.Name)
    objGridInt.colCampo.Add (ArqQtdParcEnc.Name)
    objGridInt.colCampo.Add (ArqVlrDep.Name)
    objGridInt.colCampo.Add (ArqVlrParc.Name)
    objGridInt.colCampo.Add (ArqVlrParcEnc.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_ArqSelecionado_Col = 1
    iGrid_ArqIgnorar_Col = 2
    iGrid_ArqNomeArq_Col = 3
    iGrid_ArqDataImp_Col = 4
    iGrid_ArqHoraImp_Col = 5
    iGrid_ArqQtdDep_Col = 6
    iGrid_ArqQtdParc_Col = 7
    iGrid_ArqQtdParcEnc_Col = 8
    iGrid_ArqVlrDep_Col = 9
    iGrid_ArqVlrParc_Col = 10
    iGrid_ArqVlrParcEnc_Col = 11

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridArq
    
    'Largura da primeira coluna
    GridArq.ColWidth(0) = 300

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 15

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Arq = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Arq:

    Inicializa_Grid_Arq = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202609)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Dep(objGridInt As AdmGrid) As Long
'Executa a Inicialização do griddep

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Dep

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Estabelec.")
    objGridInt.colColuna.Add ("Conta Corrente")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Bruto R$")
    objGridInt.colColuna.Add ("Comis.R$")
    objGridInt.colColuna.Add ("Líq. R$")
    objGridInt.colColuna.Add ("Sist. R$")
    objGridInt.colColuna.Add ("Parcs")
    objGridInt.colColuna.Add ("Sist.")

    'campos de edição do grid
    objGridInt.colCampo.Add (DepEstab.Name)
    objGridInt.colCampo.Add (DepCtaCorr.Name)
    objGridInt.colCampo.Add (DepDataDep.Name)
    objGridInt.colCampo.Add (DepVlrBruto.Name)
    objGridInt.colCampo.Add (DepVlrComis.Name)
    objGridInt.colCampo.Add (DepVlrLiq.Name)
    objGridInt.colCampo.Add (DepVlrParcEnc.Name)
    objGridInt.colCampo.Add (DepQtdParc.Name)
    objGridInt.colCampo.Add (DepQtdParcEnc.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_DepEstab_Col = 1
    iGrid_DepCtaCorr_Col = 2
    iGrid_DepDataDep_Col = 3
    iGrid_DepVlrBruto_Col = 4
    iGrid_DepVlrComis_Col = 5
    iGrid_DepVlrLiq_Col = 6
    iGrid_DepVlrParcEnc_Col = 7
    iGrid_DepQtdParc_Col = 8
    iGrid_DepQtdParcEnc_Col = 9

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridDep
    
    'Largura da primeira coluna
    GridDep.ColWidth(0) = 400

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 13

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Dep = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Dep:

    Inicializa_Grid_Dep = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202610)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Det(objGridInt As AdmGrid) As Long
'Executa a Inicialização do gridarq

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Det

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Compra")
    objGridInt.colColuna.Add ("Núm. Cartão")
    objGridInt.colColuna.Add ("Auto")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Bruto R$")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Núm. Título")
    objGridInt.colColuna.Add ("Sist. R$")

    'campos de edição do grid
    objGridInt.colCampo.Add (DetDataCompra.Name)
    objGridInt.colCampo.Add (DetNumCartao.Name)
    objGridInt.colCampo.Add (DetNumAuto.Name)
    objGridInt.colCampo.Add (DetParcela.Name)
    objGridInt.colCampo.Add (DetValor.Name)
    objGridInt.colCampo.Add (DetStatus.Name)
    objGridInt.colCampo.Add (DetNumTitulo.Name)
    objGridInt.colCampo.Add (DetVlrParc.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_DetDataCompra_Col = 1
    iGrid_DetNumCartao_Col = 2
    iGrid_DetNumAuto_Col = 3
    iGrid_DetParcela_Col = 4
    iGrid_DetValor_Col = 5
    iGrid_DetStatus_Col = 6
    iGrid_DetNumTitulo_Col = 7
    iGrid_DetVlrParc_Col = 8

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridDet
    
    'Largura da primeira coluna
    GridDet.ColWidth(0) = 400

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Det = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Det:

    Inicializa_Grid_Det = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202611)

    End Select

    Exit Function

End Function

Private Sub Arquivo_Click()
    If Arquivo.ListIndex <> -1 <> 0 Then Call Traz_Depositos_Tela(Arquivo.ListIndex + 1)
End Sub

Private Sub Cielo_Click()
    If Cielo.Value = vbChecked Then
        Bandeira.ListIndex = -1
        Bandeira.Enabled = False
    Else
        Bandeira.Enabled = True
        Bandeira.ListIndex = 0
    End If
End Sub

Private Sub Deposito_Click()
    Dim iExibe As Integer
    If Not gbLimpandoTela Then
        If Exibir.ListIndex <> -1 Then iExibe = Exibir.ItemData(Exibir.ListIndex)
        If Deposito.ListIndex <> -1 Then Call Traz_Detalhamento_Tela(Deposito.ListIndex + 1, iExibe)
    End If
End Sub

Private Sub Exibir_Click()
    Dim iExibe As Integer
    If Not gbLimpandoTela Then
        If Exibir.ListIndex <> -1 Then iExibe = Exibir.ItemData(Exibir.ListIndex)
        If Deposito.ListIndex <> -1 Then Call Traz_Detalhamento_Tela(Deposito.ListIndex + 1, iExibe)
    End If
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Function Trata_Parametros() As Long
    Trata_Parametros = SUCESSO
    Exit Function
End Function

Sub Limpa_Tela_BaixaCartao()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_BaixaCartao

    gbLimpandoTela = True
    
    Call Limpa_Tela(Me)

    Call Limpa_Tela_BaixaCartao_Aux
    
    Set gobjBaixaCartao = New ClassBaixaCartao
    Set gobjArq = New ClassAdmExtFinArqsLidos
    Set gobjMov = New ClassAdmExtFinMov
    Set colDet = New Collection
    
    DetTotalQtdParc.Caption = ""
    DetTotalQtdParcDif.Caption = ""
    DetTotalQtdParcEnc.Caption = ""
    DetTotalVlrParc.Caption = ""
    DetTotalVlrParcDif.Caption = ""
    DetTotalVlrParcEnc.Caption = ""
    
    DepTotalQtdParc.Caption = ""
    DepTotalQtdDep.Caption = ""
    DepTotalQtdParcEnc.Caption = ""
    DepTotalVlrParc.Caption = ""
    DepTotalVlrDep.Caption = ""
    DepTotalVlrParcEnc.Caption = ""
    
    ArqTotalQtdParc.Caption = ""
    ArqTotalQtdDep.Caption = ""
    ArqTotalQtdParcEnc.Caption = ""
    ArqTotalVlrParc.Caption = ""
    ArqTotalVlrDep.Caption = ""
    ArqTotalVlrParcEnc.Caption = ""
    
    Call Default_Tela
    
    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = TAB_TELA_SELECAO
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    Call TabStrip1_Click
    
    gbLimpandoTela = False

    Exit Sub

Erro_Limpa_Tela_BaixaCartao:

    gbLimpandoTela = False

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202612)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_BaixaCartao_Aux()
    
    'Limpa os Grids da tela
    Call Grid_Limpa(objGridArq)
    Call Grid_Limpa(objGridDep)
    Call Grid_Limpa(objGridDet)
    
    Call Ordenacao_Limpa(objGridArq)
    Call Ordenacao_Limpa(objGridDep)
    Call Ordenacao_Limpa(objGridDet)

End Sub

Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 202613

    'Limpa o restante da tela
    Call Limpa_Tela_BaixaCartao

    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 202613
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202614)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior
Dim objBaixaCartao As New ClassBaixaCartao

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    lErro = Move_Selecao_Memoria(objBaixaCartao)
    If lErro <> SUCESSO Then gError 202615

    'Se o frame anterior foi o de Seleção e ele foi alterado
    If iFrameAtual <> TAB_TELA_SELECAO And (gobjBaixaCartao.iBandeira <> objBaixaCartao.iBandeira Or _
            gobjBaixaCartao.dtDataAte <> objBaixaCartao.dtDataAte Or _
            gobjBaixaCartao.dtDataDe <> objBaixaCartao.dtDataDe) Then

        DoEvents

        lErro = Traz_BaixaCartao_Tela(objBaixaCartao)
        If lErro <> SUCESSO Then gError 202616
        
        Set gobjBaixaCartao = objBaixaCartao
        
        Call Recalcula_Lista_Arquivo

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 202615, 202616

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202617)

    End Select

    Exit Sub

End Sub

Function Move_Selecao_Memoria(ByVal objBaixaCartao As ClassBaixaCartao) As Long
'Recolhe dados do TAB de Seleção

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Move_Selecao_Memoria
   
    objBaixaCartao.iBandeira = Codigo_Extrai(Bandeira.Text)
    objBaixaCartao.iFilialEmpresa = giFilialEmpresa
    objBaixaCartao.dtDataDe = StrParaDate(DataDe.Text)
    objBaixaCartao.dtDataAte = StrParaDate(DataAte.Text)
    
    Move_Selecao_Memoria = SUCESSO

    Exit Function

Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202618)

    End Select

    Exit Function

End Function
'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object
    Set Form_Load_Ocx = Me
    Caption = "Baixa de Títulos de Cartão de Crédito"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "BaixaCartao"
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

'**** fim do trecho a ser copiado *****

Private Function Traz_BaixaCartao_Tela(ByVal objBaixaCartao As ClassBaixaCartao) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_BaixaCartao_Tela

    GL_objMDIForm.MousePointer = vbHourglass

    Call Limpa_Tela_BaixaCartao_Aux
      
    'Preenche a Coleção
    lErro = CF("BaixaCartao_Le", objBaixaCartao)
    If lErro <> SUCESSO Then gError 202619
    
    If objBaixaCartao.colArq.Count = 0 Then gError 202620
    
    lErro = Traz_Arquivos_Tela(objBaixaCartao)
    If lErro <> SUCESSO Then gError 202621
                
    GL_objMDIForm.MousePointer = vbDefault
                
    Traz_BaixaCartao_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_BaixaCartao_Tela:

    GL_objMDIForm.MousePointer = vbDefault

    Traz_BaixaCartao_Tela = gErr
    
    Select Case gErr

        Case 202619, 202621
              
        Case 202620
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_NENHUM_ARQUIVO", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202622)

    End Select

End Function

Private Function Traz_Arquivos_Tela(ByVal objBaixaCartao As ClassBaixaCartao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objArq As ClassAdmExtFinArqsLidos
Dim iQtdDep As Integer, iQtdParc As Integer, iQtdParcEnc As Integer
Dim dVlrDep As Double, dVlrParc As Double, dVlrParcEnc As Double

On Error GoTo Erro_Traz_Arquivos_Tela
       
    Call Grid_Limpa(objGridArq)
       
    If objBaixaCartao.colArq.Count >= objGridArq.objGrid.Rows Then
        Call Refaz_Grid(objGridArq, objBaixaCartao.colArq.Count)
    End If

    iIndice = 0
    For Each objArq In objBaixaCartao.colArq
    
        iIndice = iIndice + 1

        GridArq.TextMatrix(iIndice, iGrid_ArqSelecionado_Col) = CStr(MARCADO)
        GridArq.TextMatrix(iIndice, iGrid_ArqNomeArq_Col) = objArq.sNomeArq
        GridArq.TextMatrix(iIndice, iGrid_ArqDataImp_Col) = Format(objArq.dtDataImportacao, "dd/mm/yy")
        GridArq.TextMatrix(iIndice, iGrid_ArqHoraImp_Col) = Format(objArq.dHoraImportacao, "hh:mm:ss")
        
        GridArq.TextMatrix(iIndice, iGrid_ArqQtdDep_Col) = Format(objArq.iQtdDep, "#,##0")
        GridArq.TextMatrix(iIndice, iGrid_ArqQtdParc_Col) = Format(objArq.iQtdParc, "#,##0")
        GridArq.TextMatrix(iIndice, iGrid_ArqQtdParcEnc_Col) = Format(objArq.iQtdParcEnc, "#,##0")
        
        GridArq.TextMatrix(iIndice, iGrid_ArqVlrDep_Col) = Format(objArq.dVlrDep, "STANDARD")
        GridArq.TextMatrix(iIndice, iGrid_ArqVlrParc_Col) = Format(objArq.dVlrParc, "STANDARD")
        GridArq.TextMatrix(iIndice, iGrid_ArqVlrParcEnc_Col) = Format(objArq.dVlrParcEnc, "STANDARD")
        
        iQtdDep = iQtdDep + objArq.iQtdDep
        iQtdParc = iQtdParc + objArq.iQtdParc
        iQtdParcEnc = iQtdParcEnc + objArq.iQtdParcEnc
    
        dVlrDep = dVlrDep + objArq.dVlrDep
        dVlrParc = dVlrParc + objArq.dVlrParc
        dVlrParcEnc = dVlrParcEnc + objArq.dVlrParcEnc
        
    Next
    
    objGridArq.iLinhasExistentes = objBaixaCartao.colArq.Count
   
    Call Grid_Refresh_Checkbox(objGridArq)
    
    ArqTotalQtdParc.Caption = Format(iQtdParc, "#,##0")
    ArqTotalQtdDep.Caption = Format(iQtdDep, "#,##0")
    ArqTotalQtdParcEnc.Caption = Format(iQtdParcEnc, "#,##0")
    
    ArqTotalVlrParc.Caption = Format(dVlrParc, "STANDARD")
    ArqTotalVlrDep.Caption = Format(dVlrDep, "STANDARD")
    ArqTotalVlrParcEnc.Caption = Format(dVlrParcEnc, "STANDARD")
                
    Traz_Arquivos_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Arquivos_Tela:

    Traz_Arquivos_Tela = gErr
    
    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202623)

    End Select

End Function

Private Sub Recalcula_Lista_Arquivo()

Dim objArq As ClassAdmExtFinArqsLidos
Dim iIndice As Integer

On Error GoTo Erro_Recalcula_Lista_Arquivo

    Arquivo.Clear
    For Each objArq In gobjBaixaCartao.colArq
        iIndice = iIndice + 1
        Arquivo.AddItem "Item " & CStr(iIndice) & SEPARADOR & objArq.sNomeArq
    Next
    If Arquivo.ListCount > 0 Then
        Arquivo.ListIndex = 0
        Call Arquivo_Click
    End If
    
    Exit Sub

Erro_Recalcula_Lista_Arquivo:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202624)

    End Select

    Exit Sub
    
End Sub

Private Sub Recalcula_Lista_Deposito()

Dim objMov As ClassAdmExtFinMov
Dim iIndice As Integer

On Error GoTo Erro_Recalcula_Lista_Deposito

    Deposito.Clear
    For Each objMov In gobjArq.colMov
        iIndice = iIndice + 1
        Deposito.AddItem "Item " & CStr(iIndice) & SEPARADOR & "Estab.: " & objMov.sEstabelecimento & " Data: " & Format(objMov.dtData, "dd/mm/yyyy") & " e Valor: " & Format(objMov.dValorBruto, "STANDARD")
    Next
    If Deposito.ListCount > 0 Then
        Deposito.ListIndex = 0
        Call Deposito_Click
    End If
    
    Exit Sub

Erro_Recalcula_Lista_Deposito:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202625)

    End Select

    Exit Sub
    
End Sub

Private Function Traz_Depositos_Tela(ByVal iIndiceArq As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objArq As ClassAdmExtFinArqsLidos
Dim objMov As ClassAdmExtFinMov
Dim iQtdDep As Integer, iQtdParc As Integer, iQtdParcEnc As Integer
Dim dVlrDep As Double, dVlrParc As Double, dVlrParcEnc As Double

On Error GoTo Erro_Traz_Depositos_Tela

    Call Grid_Limpa(objGridDep)

    Set objArq = gobjBaixaCartao.colArq.Item(iIndiceArq)
       
    If objArq.colMov.Count >= objGridDep.objGrid.Rows Then
        Call Refaz_Grid(objGridDep, objArq.colMov.Count)
    End If

    iIndice = 0
    For Each objMov In objArq.colMov
    
        iIndice = iIndice + 1

        GridDep.TextMatrix(iIndice, iGrid_DepEstab_Col) = objMov.sEstabelecimento
        GridDep.TextMatrix(iIndice, iGrid_DepCtaCorr_Col) = objMov.iCodConta & SEPARADOR & objMov.sNomeCtaCorrente
        GridDep.TextMatrix(iIndice, iGrid_DepDataDep_Col) = Format(objMov.dtData, "dd/mm/yy")
        GridDep.TextMatrix(iIndice, iGrid_DepVlrBruto_Col) = Format(objMov.dValorBruto, "STANDARD")
        GridDep.TextMatrix(iIndice, iGrid_DepVlrComis_Col) = Format(objMov.dValorComissao, "STANDARD")
        GridDep.TextMatrix(iIndice, iGrid_DepVlrLiq_Col) = Format(objMov.dValorLiq, "STANDARD")
        
        GridDep.TextMatrix(iIndice, iGrid_DepQtdParc_Col) = Format(objMov.iQtdParc, "#,##0")
        GridDep.TextMatrix(iIndice, iGrid_DepQtdParcEnc_Col) = Format(objMov.iQtdParcEnc, "#,##0")
        
        GridDep.TextMatrix(iIndice, iGrid_DepVlrParcEnc_Col) = Format(objMov.dVlrParcEnc, "STANDARD")
        
        iQtdDep = iQtdDep + objMov.iQtdDep
        iQtdParc = iQtdParc + objMov.iQtdParc
        iQtdParcEnc = iQtdParcEnc + objMov.iQtdParcEnc
    
        dVlrDep = dVlrDep + objMov.dVlrDep
        dVlrParc = dVlrParc + objMov.dVlrParc
        dVlrParcEnc = dVlrParcEnc + objMov.dVlrParcEnc
                
    Next
    
    objGridDep.iLinhasExistentes = objArq.colMov.Count

    DepTotalQtdParc.Caption = Format(iQtdParc, "#,##0")
    DepTotalQtdDep.Caption = Format(iQtdDep, "#,##0")
    DepTotalQtdParcEnc.Caption = Format(iQtdParcEnc, "#,##0")
    
    DepTotalVlrParc.Caption = Format(dVlrParc, "STANDARD")
    DepTotalVlrDep.Caption = Format(dVlrDep, "STANDARD")
    DepTotalVlrParcEnc.Caption = Format(dVlrParcEnc, "STANDARD")
            
    Set gobjArq = objArq
    
    Call Recalcula_Lista_Deposito
                
    Traz_Depositos_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Depositos_Tela:

    Traz_Depositos_Tela = gErr
    
    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202626)

    End Select

End Function

Private Function Traz_Detalhamento_Tela(ByVal iIndiceDep As Integer, Optional ByVal iExibe As Integer = TELA_EXIBE_TODOS) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objMov As ClassAdmExtFinMov
Dim objDet As ClassAdmExtFinMovDet
Dim iQtdParc As Integer, iQtdParcEnc As Integer
Dim dVlrParc As Double, dVlrParcEnc As Double

On Error GoTo Erro_Traz_Detalhamento_Tela

    Call Grid_Limpa(objGridDet)

    Set objMov = gobjArq.colMov.Item(iIndiceDep)
       
    If objMov.colMovDet.Count >= objGridDet.objGrid.Rows Then
        Call Refaz_Grid(objGridDet, objMov.colMovDet.Count)
    End If

    iIndice = 0
    Set colDet = New Collection
    For Each objDet In objMov.colMovDet
    
        If iExibe = TELA_EXIBE_TODOS Or (iExibe = TELA_COM_ERROS And objDet.iCodErro > 0) Or (iExibe = TELA_SEM_ERROS And objDet.iCodErro = 0) Then
    
            iIndice = iIndice + 1
    
            GridDet.TextMatrix(iIndice, iGrid_DetDataCompra_Col) = Format(objDet.dtDataCompra, "dd/mm/yy")
            GridDet.TextMatrix(iIndice, iGrid_DetNumCartao_Col) = objDet.sNumCartao
            GridDet.TextMatrix(iIndice, iGrid_DetNumAuto_Col) = objDet.sAutorizacao
            GridDet.TextMatrix(iIndice, iGrid_DetParcela_Col) = CStr(objDet.iNumParcela)
            GridDet.TextMatrix(iIndice, iGrid_DetStatus_Col) = CStr(objDet.iCodErro)
            GridDet.TextMatrix(iIndice, iGrid_DetNumTitulo_Col) = CStr(objDet.lNumTitulo)
            GridDet.TextMatrix(iIndice, iGrid_DetVlrParc_Col) = Format(objDet.dVlrParcEnc, "STANDARD")
            GridDet.TextMatrix(iIndice, iGrid_DetValor_Col) = Format(objDet.dValor, "STANDARD")
            
            iQtdParc = iQtdParc + objDet.iQtdParc
            iQtdParcEnc = iQtdParcEnc + objDet.iQtdParcEnc
        
            dVlrParc = dVlrParc + objDet.dVlrParc
            dVlrParcEnc = dVlrParcEnc + objDet.dVlrParcEnc
            
            colDet.Add objDet
            
        End If
        
    Next
    
    objGridDet.iLinhasExistentes = iIndice

    DetTotalQtdParc.Caption = Format(iQtdParc, "#,##0")
    DetTotalQtdParcDif.Caption = Format(iQtdParc - iQtdParcEnc, "#,##0")
    DetTotalQtdParcEnc.Caption = Format(iQtdParcEnc, "#,##0")
    
    DetTotalVlrParc.Caption = Format(dVlrParc, "STANDARD")
    DetTotalVlrParcDif.Caption = Format(dVlrParc - dVlrParcEnc, "STANDARD")
    DetTotalVlrParcEnc.Caption = Format(dVlrParcEnc, "STANDARD")
            
    Set gobjMov = objMov
    
    If objMov.colMovDet.Count > 0 Then Call Trata_Status(StrParaInt(GridDet.TextMatrix(1, iGrid_DetStatus_Col)))
            
    Traz_Detalhamento_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Detalhamento_Tela:

    Traz_Detalhamento_Tela = gErr
    
    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202627)

    End Select

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    
    objGridInt.objGrid.Rows = iNumLinhas + 1
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
End Sub

Private Sub GridArq_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection
Dim iLinha As Integer

    Call Grid_Click(objGridArq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridArq, iAlterado)
    End If
    
    colcolColecao.Add gobjBaixaCartao.colArq
    
    Call Ordenacao_ClickGrid(objGridArq, , colcolColecao)
    
    Call Recalcula_Lista_Arquivo

End Sub

Private Sub GridArq_GotFocus()
    Call Grid_Recebe_Foco(objGridArq)
End Sub

Private Sub GridArq_EnterCell()
    Call Grid_Entrada_Celula(objGridArq, iAlterado)
End Sub

Private Sub GridArq_LeaveCell()
    Call Saida_Celula(objGridArq)
End Sub

Private Sub GridArq_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridArq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridArq, iAlterado)
    End If

End Sub

Private Sub GridArq_RowColChange()
    Call Grid_RowColChange(objGridArq)
End Sub

Private Sub GridArq_Scroll()
    Call Grid_Scroll(objGridArq)
End Sub

Private Sub GridArq_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridArq)

End Sub

Private Sub GridArq_LostFocus()
    Call Grid_Libera_Foco(objGridArq)
End Sub

Private Sub GridDep_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection
Dim iLinha As Integer

    Call Grid_Click(objGridDep, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDep, iAlterado)
    End If
    
    colcolColecao.Add gobjArq.colMov
    
    Call Ordenacao_ClickGrid(objGridDep, , colcolColecao)
    
    Call Recalcula_Lista_Deposito

End Sub

Private Sub GridDep_GotFocus()
    Call Grid_Recebe_Foco(objGridDep)
End Sub

Private Sub GridDep_EnterCell()
    Call Grid_Entrada_Celula(objGridDep, iAlterado)
End Sub

Private Sub GridDep_LeaveCell()
    Call Saida_Celula(objGridDep)
End Sub

Private Sub GridDep_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDep, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDep, iAlterado)
    End If

End Sub

Private Sub GridDep_RowColChange()
    Call Grid_RowColChange(objGridDep)
End Sub

Private Sub GridDep_Scroll()
    Call Grid_Scroll(objGridDep)
End Sub

Private Sub GridDep_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridDep)
End Sub

Private Sub GridDep_LostFocus()
    Call Grid_Libera_Foco(objGridDep)
End Sub

Private Sub GridDet_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection
Dim iLinha As Integer

    Call Grid_Click(objGridDet, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDet, iAlterado)
    End If
    
    colcolColecao.Add gobjMov.colMovDet
    colcolColecao.Add colDet
    
    Call Ordenacao_ClickGrid(objGridDet, , colcolColecao)

End Sub

Private Sub GridDet_GotFocus()
    Call Grid_Recebe_Foco(objGridDet)
End Sub

Private Sub GridDet_EnterCell()
    Call Grid_Entrada_Celula(objGridDet, iAlterado)
End Sub

Private Sub GridDet_LeaveCell()
    Call Saida_Celula(objGridDet)
End Sub

Private Sub GridDet_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDet, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDet, iAlterado)
    End If

End Sub

Private Sub GridDet_RowColChange()
    Call Grid_RowColChange(objGridDet)
    
    If objGridDet.iLinhaAntiga <> objGridDet.objGrid.Row And GridDet.Row <> 0 Then
        Call Trata_Status(StrParaInt(GridDet.TextMatrix(GridDet.Row, iGrid_DetStatus_Col)))
        objGridDet.iLinhaAntiga = objGridDet.objGrid.Row
    End If
End Sub

Private Sub GridDet_Scroll()
    Call Grid_Scroll(objGridDet)
End Sub

Private Sub GridDet_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridDet)
End Sub

Private Sub GridDet_LostFocus()
    Call Grid_Libera_Foco(objGridDet)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridArq.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_ArqSelecionado_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, ArqSelecionado)
                    If lErro <> SUCESSO Then gError 202628
                     
                Case iGrid_ArqIgnorar_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, ArqIgnorar)
                    If lErro <> SUCESSO Then gError 202628
                     
            End Select
            
 
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 202629

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 202628

        Case 202629
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202630)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
    
        Case ArqSelecionado.Name
            objControl.Enabled = True
            
        Case ArqIgnorar.Name
            If StrParaInt(GridArq.TextMatrix(iLinha, iGrid_ArqSelecionado_Col)) = MARCADO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        Case Else
            objControl.Enabled = False
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202631)

    End Select

    Exit Sub

End Sub

Public Sub ArqSelecionado_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Calcula_TotalArq
End Sub

Public Sub ArqSelecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArq)
End Sub

Public Sub ArqSelecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArq)
End Sub

Public Sub ArqSelecionado_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridArq.objControle = ArqSelecionado
    lErro = Grid_Campo_Libera_Foco(objGridArq)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ArqIgnorar_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Calcula_TotalArq
End Sub

Public Sub ArqIgnorar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArq)
End Sub

Public Sub ArqIgnorar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArq)
End Sub

Public Sub ArqIgnorar_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridArq.objControle = ArqIgnorar
    lErro = Grid_Campo_Libera_Foco(objGridArq)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub Marca_Desmarca(ByVal iMarcado As Integer)
'Marca todos os bloqueios do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridArq.iLinhasExistentes
        GridArq.TextMatrix(iLinha, iGrid_ArqSelecionado_Col) = CStr(iMarcado)
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridArq)
    
    Call Calcula_TotalArq
    
End Sub

Private Sub BotaoMarcarTodos_Click()
    Call Marca_Desmarca(MARCADO)
End Sub

Private Sub BotaoDesmarcarTodos_Click(Index As Integer)
    Call Marca_Desmarca(DESMARCADO)
End Sub

Sub Default_Tela()

Dim lErro As Long

On Error GoTo Erro_Default_Tela

    Bandeira.ListIndex = 0
    Exibir.ListIndex = 0
    
    Exit Sub

Erro_Default_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202632)

    End Select

    Exit Sub

End Sub

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 202633

    'Limpa Tela
    Call Limpa_Tela_BaixaCartao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 202633

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202634)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objBaixaCartao As New ClassBaixaCartao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Preenche o objBaixaCartao
    lErro = Move_Tela_Memoria(objBaixaCartao)
    If lErro <> SUCESSO Then gError 202635
    
    If objBaixaCartao.colArq.Count = 0 Then gError 202636

    lErro = CF("AdmExtFin_AtualizarExtratos", objBaixaCartao)
    If lErro <> SUCESSO Then gError 202637

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 202635, 202637
        
        Case 202636
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ARQUIVO_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202638)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal objBaixaCartao As ClassBaixaCartao) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objArq As ClassAdmExtFinArqsLidos

On Error GoTo Erro_Move_Tela_Memoria

    iLinha = 0
    For Each objArq In gobjBaixaCartao.colArq
        iLinha = iLinha + 1
        If StrParaInt(GridArq.TextMatrix(iLinha, iGrid_ArqSelecionado_Col)) = MARCADO Then
            objArq.iNaoAtualizar = StrParaInt(GridArq.TextMatrix(iLinha, iGrid_ArqIgnorar_Col))
            objBaixaCartao.colArq.Add objArq
            objBaixaCartao.iTotalReg = objBaixaCartao.iTotalReg + objArq.iTotalReg
        End If
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202639)

    End Select

End Function

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_DownClick

    DataDe.SetFocus

    If Len(DataDe.ClipText) > 0 Then

        sData = DataDe.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 202640

        DataDe.Text = sData

    End If

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 202640

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202641)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_UpClick

    DataDe.SetFocus

    If Len(Trim(DataDe.ClipText)) > 0 Then

        sData = DataDe.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 202642

        DataDe.Text = sData

    End If

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 202642

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202643)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(Trim(DataDe.ClipText)) <> 0 Then

        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 202644

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 202644

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202645)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_DownClick

    DataAte.SetFocus

    If Len(DataAte.ClipText) > 0 Then

        sData = DataAte.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 202646

        DataAte.Text = sData

    End If

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 202646

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202647)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_UpClick

    DataAte.SetFocus

    If Len(Trim(DataAte.ClipText)) > 0 Then

        sData = DataAte.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 202648

        DataAte.Text = sData

    End If

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 202648

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202649)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(Trim(DataAte.ClipText)) <> 0 Then

        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 202650

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 202650

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202651)

    End Select

    Exit Sub

End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Calcula_TotalArq()

Dim lErro As Long
Dim iLinha As Integer
Dim objArq As ClassAdmExtFinArqsLidos
Dim iQtdDep As Integer, iQtdParc As Integer, iQtdParcEnc As Integer
Dim dVlrDep As Double, dVlrParc As Double, dVlrParcEnc As Double

On Error GoTo Erro_Calcula_TotalArq

    iLinha = 0
    For Each objArq In gobjBaixaCartao.colArq
        iLinha = iLinha + 1
        If StrParaInt(GridArq.TextMatrix(iLinha, iGrid_ArqSelecionado_Col)) = MARCADO Then
            iQtdDep = iQtdDep + objArq.iQtdDep
            iQtdParc = iQtdParc + objArq.iQtdParc
            iQtdParcEnc = iQtdParcEnc + objArq.iQtdParcEnc
        
            dVlrDep = dVlrDep + objArq.dVlrDep
            dVlrParc = dVlrParc + objArq.dVlrParc
            dVlrParcEnc = dVlrParcEnc + objArq.dVlrParcEnc
        End If
    Next
    
    ArqTotalQtdParc.Caption = Format(iQtdParc, "#,##0")
    ArqTotalQtdDep.Caption = Format(iQtdDep, "#,##0")
    ArqTotalQtdParcEnc.Caption = Format(iQtdParcEnc, "#,##0")
    
    ArqTotalVlrParc.Caption = Format(dVlrParc, "STANDARD")
    ArqTotalVlrDep.Caption = Format(dVlrDep, "STANDARD")
    ArqTotalVlrParcEnc.Caption = Format(dVlrParcEnc, "STANDARD")

    Exit Sub

Erro_Calcula_TotalArq:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202652)

    End Select

    Exit Sub

End Sub

Sub Trata_Status(ByVal iCodErro As Integer)

    Select Case iCodErro
    
        Case ADMEXTFIN_ERRO_CV_DIFNUMCARTAO
            Status.Text = "A parcela foi encontrada somente para outro número de cartão."
        
        Case ADMEXTFIN_ERRO_CV_MUITASPARCELAS
            Status.Text = "Várias parcelas atendem aos critérios."

        Case ADMEXTFIN_ERRO_CV_PARCNAOABERTA
            Status.Text = "A parcela já está baixada."

        Case ADMEXTFIN_ERRO_CV_SEMPARCELA
            Status.Text = "Nenhuma parcela encontrada."
            
        Case 0
            Status.Text = "Parcela localizada com sucesso."
            
    End Select
    
End Sub

Private Sub BotaoTitulo_Click()

Dim lErro As Long
Dim objTitRec As New ClassTituloReceber

On Error GoTo Erro_BotaoTitulo

    If GridDet.Row = 0 Then gError 202653
    
    objTitRec.iFilialEmpresa = giFilialEmpresa
    'objTitRec.lNumIntDoc = gobjMov.colMovDet.Item(GridDet.Row).lNumIntTitulo
    objTitRec.lNumIntDoc = colDet.Item(GridDet.Row).lNumIntTitulo

    Call Chama_Tela("TituloReceber_Consulta", objTitRec)

    Exit Sub
    
Erro_BotaoTitulo:
    
    Select Case gErr
    
        Case 202653
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202654)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExibeParcCartao_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sNomeBrowse As String
Dim sFiltro As String

On Error GoTo Erro_BotaoExibeParcCartao_Click

    'If GridDet.Row = 0 Then gError 202655
    
    lErro = CF("VendasComCartao_Obtem_NomeBrowser", sNomeBrowse, sFiltro, colSelecao, VENDCARTAO_BROWSER_TIPO_TODASABERTAS, GridDet.TextMatrix(GridDet.Row, iGrid_DetNumCartao_Col), StrParaDate(GridDet.TextMatrix(GridDet.Row, iGrid_DetDataCompra_Col)), StrParaDbl(GridDet.TextMatrix(GridDet.Row, iGrid_DetValor_Col)), GridDet.TextMatrix(GridDet.Row, iGrid_DetNumAuto_Col))
    If lErro <> SUCESSO Then gError 202656
    
    Call Chama_Tela(sNomeBrowse, colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoExibeParcCartao_Click:
    
    Select Case gErr
    
        Case 202655
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 202656

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202657)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExibeParcDataValor_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sNomeBrowse As String
Dim sFiltro As String

On Error GoTo Erro_BotaoExibeParcDataValor_Click

    If GridDet.Row = 0 Then gError 202658
    
    lErro = CF("VendasComCartao_Obtem_NomeBrowser", sNomeBrowse, sFiltro, colSelecao, VENDCARTAO_BROWSER_TIPO_DATAVALOR, GridDet.TextMatrix(GridDet.Row, iGrid_DetNumCartao_Col), StrParaDate(GridDet.TextMatrix(GridDet.Row, iGrid_DetDataCompra_Col)), StrParaDbl(GridDet.TextMatrix(GridDet.Row, iGrid_DetValor_Col)), GridDet.TextMatrix(GridDet.Row, iGrid_DetNumAuto_Col))
    If lErro <> SUCESSO Then gError 202659
    
    Call Chama_Tela(sNomeBrowse, colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoExibeParcDataValor_Click:
    
    Select Case gErr
    
        Case 202658
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 202659

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202660)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExibeParcEsseCartao_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sNomeBrowse As String
Dim sFiltro As String

On Error GoTo Erro_BotaoExibeParcEsseCartao_Click

    If GridDet.Row = 0 Then gError 202661
    
    lErro = CF("VendasComCartao_Obtem_NomeBrowser", sNomeBrowse, sFiltro, colSelecao, VENDCARTAO_BROWSER_TIPO_DESSECARTAO, GridDet.TextMatrix(GridDet.Row, iGrid_DetNumCartao_Col), StrParaDate(GridDet.TextMatrix(GridDet.Row, iGrid_DetDataCompra_Col)), StrParaDbl(GridDet.TextMatrix(GridDet.Row, iGrid_DetValor_Col)), GridDet.TextMatrix(GridDet.Row, iGrid_DetNumAuto_Col))
    If lErro <> SUCESSO Then gError 202662
    
    Call Chama_Tela(sNomeBrowse, colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoExibeParcEsseCartao_Click:
    
    Select Case gErr
    
        Case 202661
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 202662

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202663)
    
    End Select
    
    Exit Sub
    
End Sub

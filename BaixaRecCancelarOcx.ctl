VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaRecCancelarOcx 
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   2
      Left            =   75
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   9285
      Begin VB.Frame FrameDetalhamentoBaixa 
         BorderStyle     =   0  'None
         Height          =   2420
         Index           =   2
         Left            =   630
         TabIndex        =   80
         Top             =   2880
         Visible         =   0   'False
         Width           =   8175
         Begin VB.Frame FrameDadosBaixa 
            Caption         =   "Dados da Baixa"
            Height          =   2415
            Left            =   15
            TabIndex        =   86
            Top             =   -15
            Width           =   8175
            Begin VB.Frame FrameRecebimento 
               Caption         =   "Cartão"
               Height          =   1260
               Index           =   10
               Left            =   120
               TabIndex        =   186
               Top             =   1080
               Visible         =   0   'False
               Width           =   7950
               Begin VB.Label Label47 
                  Caption         =   "Data Emissão:"
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
                  Left            =   5070
                  TabIndex        =   196
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.Label Label51 
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
                  Height          =   255
                  Left            =   345
                  TabIndex        =   195
                  Top             =   810
                  Width           =   450
               End
               Begin VB.Label Label55 
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
                  Height          =   195
                  Left            =   2940
                  TabIndex        =   194
                  Top             =   345
                  Width           =   465
               End
               Begin VB.Label DataEmissaoTitCartao 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   6405
                  TabIndex        =   193
                  Top             =   300
                  Width           =   1095
               End
               Begin VB.Label TipoTitCartao 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   885
                  TabIndex        =   192
                  Top             =   780
                  Width           =   1080
               End
               Begin VB.Label FilialTitCartao 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   3510
                  TabIndex        =   191
                  Top             =   315
                  Width           =   1095
               End
               Begin VB.Label NumTitCartao 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   3510
                  TabIndex        =   190
                  Top             =   780
                  Width           =   720
               End
               Begin VB.Label ClienteTitCartao 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   885
                  TabIndex        =   189
                  Top             =   315
                  Width           =   1740
               End
               Begin VB.Label Label52 
                  Caption         =   "Número:"
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
                  Left            =   2700
                  TabIndex        =   188
                  Top             =   810
                  Width           =   705
               End
               Begin VB.Label Label54 
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
                  Height          =   195
                  Left            =   135
                  TabIndex        =   187
                  Top             =   345
                  Width           =   660
               End
            End
            Begin VB.Frame FrameRecebimento 
               Caption         =   "Dados do Recebimento"
               Height          =   1260
               Index           =   7
               Left            =   120
               TabIndex        =   110
               Top             =   1080
               Width           =   7950
               Begin VB.Label Historico 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   720
                  Left            =   2940
                  TabIndex        =   116
                  Top             =   430
                  Width           =   4800
               End
               Begin VB.Label Valor 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   975
                  TabIndex        =   115
                  Top             =   780
                  Width           =   1590
               End
               Begin VB.Label ContaCorrente 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   960
                  TabIndex        =   114
                  Top             =   300
                  Width           =   1590
               End
               Begin VB.Label Label25 
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
                  Height          =   255
                  Left            =   375
                  TabIndex        =   113
                  Top             =   810
                  Width           =   495
               End
               Begin VB.Label Label13 
                  Caption         =   "Histórico:"
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
                  Left            =   2925
                  TabIndex        =   112
                  Top             =   180
                  Width           =   810
               End
               Begin VB.Label Label4 
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
                  Left            =   120
                  TabIndex        =   111
                  Top             =   360
                  Width           =   555
               End
            End
            Begin VB.Frame FrameRecebimento 
               Caption         =   "Perda"
               Height          =   1260
               Index           =   4
               Left            =   120
               TabIndex        =   139
               Top             =   1080
               Visible         =   0   'False
               Width           =   7950
               Begin VB.Label Label16 
                  Caption         =   "Histórico:"
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
                  Left            =   630
                  TabIndex        =   141
                  Top             =   360
                  Width           =   810
               End
               Begin VB.Label HistoricoPerda 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   630
                  TabIndex        =   140
                  Top             =   600
                  Width           =   6630
               End
            End
            Begin VB.Frame FrameRecebimento 
               Caption         =   "Adiantamento de Cliente"
               Height          =   1260
               Index           =   5
               Left            =   120
               TabIndex        =   117
               Top             =   1080
               Visible         =   0   'False
               Width           =   7950
               Begin VB.Label CCIntNomeReduzido 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "CCorrente"
                  Height          =   300
                  Left            =   5160
                  TabIndex        =   125
                  Top             =   255
                  Width           =   1095
               End
               Begin VB.Label ValorPA 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "ValorPagtoAnt"
                  Height          =   300
                  Left            =   5160
                  TabIndex        =   124
                  Top             =   765
                  Width           =   1095
               End
               Begin VB.Label MeioPagtoDescricao 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "MeioPagto"
                  Height          =   300
                  Left            =   2070
                  TabIndex        =   123
                  Top             =   780
                  Width           =   1095
               End
               Begin VB.Label DataMovimento 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "DataMovto"
                  Height          =   300
                  Left            =   2070
                  TabIndex        =   122
                  Top             =   270
                  Width           =   1095
               End
               Begin VB.Label Label23 
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
                  Height          =   255
                  Left            =   4560
                  TabIndex        =   121
                  Top             =   795
                  Width           =   510
               End
               Begin VB.Label Label7 
                  Caption         =   "Meio Pagto:"
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
                  Left            =   900
                  TabIndex        =   120
                  Top             =   840
                  Width           =   1035
               End
               Begin VB.Label Label6 
                  Caption         =   "C/C.:"
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
                  Left            =   4590
                  TabIndex        =   119
                  Top             =   300
                  Width           =   480
               End
               Begin VB.Label Label5 
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
                  Height          =   255
                  Left            =   1455
                  TabIndex        =   118
                  Top             =   315
                  Width           =   540
               End
            End
            Begin VB.Frame FrameRecebimento 
               Caption         =   "Débito"
               Height          =   1260
               Index           =   6
               Left            =   120
               TabIndex        =   126
               Top             =   1080
               Visible         =   0   'False
               Width           =   7950
               Begin VB.Label SaldoDebito 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   4020
                  TabIndex        =   138
                  Top             =   780
                  Width           =   1080
               End
               Begin VB.Label NumTitulo 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   4020
                  TabIndex        =   137
                  Top             =   270
                  Width           =   720
               End
               Begin VB.Label FilialEmpresaCR 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "FilEmpr"
                  Height          =   300
                  Left            =   -20000
                  TabIndex        =   136
                  Top             =   1192
                  Visible         =   0   'False
                  Width           =   525
               End
               Begin VB.Label ValorDebito 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1410
                  TabIndex        =   135
                  Top             =   780
                  Width           =   1080
               End
               Begin VB.Label SiglaDocumentoCR 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   6405
                  TabIndex        =   134
                  Top             =   270
                  Width           =   1080
               End
               Begin VB.Label DataEmissaoCred 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1395
                  TabIndex        =   133
                  Top             =   270
                  Width           =   1095
               End
               Begin VB.Label Label40 
                  Caption         =   "Filial Empresa:"
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
                  Left            =   -20000
                  TabIndex        =   132
                  Top             =   1215
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.Label Label39 
                  AutoSize        =   -1  'True
                  Caption         =   "Saldo:"
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
                  Left            =   3330
                  TabIndex        =   131
                  Top             =   810
                  Width           =   555
               End
               Begin VB.Label Label38 
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
                  Left            =   765
                  TabIndex        =   130
                  Top             =   825
                  Width           =   510
               End
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "Número:"
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
                  Left            =   3165
                  TabIndex        =   129
                  Top             =   315
                  Width           =   720
               End
               Begin VB.Label Label48 
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
                  Left            =   5835
                  TabIndex        =   128
                  Top             =   315
                  Width           =   450
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  Caption         =   "Emissão:"
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
                  Left            =   510
                  TabIndex        =   127
                  Top             =   315
                  Width           =   765
               End
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Valor Baixado:"
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
               Left            =   5400
               TabIndex        =   98
               Top             =   330
               Width           =   1245
            End
            Begin VB.Label ValorBaixado 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   6705
               TabIndex        =   97
               Top             =   270
               Width           =   945
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Desconto:"
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
               TabIndex        =   96
               Top             =   735
               Width           =   885
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Multa:"
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
               Left            =   3405
               TabIndex        =   95
               Top             =   735
               Width           =   540
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Juros:"
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
               Left            =   6120
               TabIndex        =   94
               Top             =   735
               Width           =   525
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Valor Pago:"
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
               Left            =   2940
               TabIndex        =   93
               Top             =   330
               Width           =   1005
            End
            Begin VB.Label Desconto 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1350
               TabIndex        =   92
               Top             =   675
               Width           =   945
            End
            Begin VB.Label Multa 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3990
               TabIndex        =   91
               Top             =   675
               Width           =   945
            End
            Begin VB.Label Juros 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   6705
               TabIndex        =   90
               Top             =   675
               Width           =   945
            End
            Begin VB.Label ValorPago 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   4005
               TabIndex        =   89
               Top             =   270
               Width           =   945
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Data Baixa:"
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
               TabIndex        =   88
               Top             =   285
               Width           =   1005
            End
            Begin VB.Label DataBaixa 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1350
               TabIndex        =   87
               Top             =   240
               Width           =   1125
            End
         End
         Begin VB.Frame FrameBaixasMovCCI 
            Caption         =   "Baixas do Movimento Selecionado"
            Height          =   2415
            Left            =   15
            TabIndex        =   100
            Top             =   -15
            Visible         =   0   'False
            Width           =   8100
            Begin MSMask.MaskEdBox SequencialBaixada 
               Height          =   225
               Left            =   3480
               TabIndex        =   101
               Top             =   1080
               Width           =   1020
               _ExtentX        =   1799
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
            Begin MSMask.MaskEdBox ValorParcelaBaixada 
               Height          =   225
               Left            =   2160
               TabIndex        =   102
               Top             =   1080
               Width           =   1260
               _ExtentX        =   2223
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
            Begin MSMask.MaskEdBox ValorBaixada 
               Height          =   225
               Left            =   4800
               TabIndex        =   103
               Top             =   1080
               Width           =   1260
               _ExtentX        =   2223
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
            Begin MSMask.MaskEdBox ParcelaBaixada 
               Height          =   225
               Left            =   840
               TabIndex        =   104
               Top             =   1080
               Width           =   900
               _ExtentX        =   1588
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
            Begin MSMask.MaskEdBox NumeroBaixada 
               Height          =   225
               Left            =   2160
               TabIndex        =   105
               Top             =   720
               Width           =   1020
               _ExtentX        =   1799
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
            Begin MSMask.MaskEdBox TipoBaixada 
               Height          =   225
               Left            =   4800
               TabIndex        =   106
               Top             =   720
               Width           =   780
               _ExtentX        =   1376
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
            Begin MSMask.MaskEdBox FilialEmpresaBaixada 
               Height          =   225
               Left            =   3480
               TabIndex        =   107
               Top             =   720
               Width           =   1260
               _ExtentX        =   2223
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
            Begin MSFlexGridLib.MSFlexGrid GridBaixasMovimento 
               Height          =   2055
               Left            =   240
               TabIndex        =   108
               Top             =   240
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   3625
               _Version        =   393216
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Tipo de Baixa"
            Height          =   1290
            Left            =   -10000
            TabIndex        =   81
            Top             =   1440
            Visible         =   0   'False
            Width           =   2535
            Begin VB.OptionButton Recebimento 
               Caption         =   "Recebimento"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   120
               TabIndex        =   85
               TabStop         =   0   'False
               Top             =   225
               Value           =   -1  'True
               Width           =   1995
            End
            Begin VB.OptionButton Recebimento 
               Caption         =   "Adiantamento"
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   135
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   465
               Width           =   2145
            End
            Begin VB.OptionButton Recebimento 
               Caption         =   "Débito / Devolução"
               Enabled         =   0   'False
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
               Left            =   120
               TabIndex        =   83
               TabStop         =   0   'False
               Top             =   750
               Width           =   2115
            End
            Begin VB.OptionButton Recebimento 
               Caption         =   "Perda"
               Enabled         =   0   'False
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
               Left            =   120
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   1005
               Width           =   2115
            End
         End
      End
      Begin VB.Frame FrameDetalhamentoBaixa 
         Caption         =   "Cancelamento"
         Height          =   2415
         Index           =   1
         Left            =   630
         TabIndex        =   70
         Top             =   2850
         Width           =   8175
         Begin VB.TextBox HistoricoCancelamento 
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   71
            Top             =   1800
            Width           =   6135
         End
         Begin MSComCtl2.UpDown UpDownDataCancelamento 
            Height          =   300
            Left            =   2760
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   420
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataCancelamento 
            Height          =   300
            Left            =   1680
            TabIndex        =   73
            Top             =   420
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label TotalCancelar 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1680
            TabIndex        =   79
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label ItensSelecionados 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5880
            TabIndex        =   78
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Total a cancelar:"
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
            TabIndex        =   77
            Top             =   1230
            Width           =   1470
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Itens Selecionados:"
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
            Left            =   4080
            TabIndex        =   76
            Top             =   1245
            Width           =   1695
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cancelar em:"
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
            Left            =   540
            TabIndex        =   75
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Histórico:"
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
            Left            =   840
            TabIndex        =   74
            Top             =   1860
            Width           =   825
         End
      End
      Begin VB.Frame FrameBaixaBaixa 
         Caption         =   "Parcelas Baixadas"
         Height          =   2415
         Left            =   480
         TabIndex        =   8
         Top             =   0
         Width           =   8415
         Begin MSMask.MaskEdBox FilialEmpresa 
            Height          =   240
            Left            =   3240
            TabIndex        =   56
            Top             =   360
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   423
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
         Begin MSMask.MaskEdBox ValorPagoBaixa 
            Height          =   225
            Left            =   3960
            TabIndex        =   55
            Top             =   1470
            Width           =   1260
            _ExtentX        =   2223
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
         Begin MSMask.MaskEdBox Tipo 
            Height          =   225
            Left            =   240
            TabIndex        =   10
            Top             =   1440
            Width           =   615
            _ExtentX        =   1085
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
         Begin MSMask.MaskEdBox Numero 
            Height          =   225
            Left            =   1005
            TabIndex        =   9
            Top             =   1680
            Width           =   780
            _ExtentX        =   1376
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
         Begin MSMask.MaskEdBox Parcela 
            Height          =   225
            Left            =   1875
            TabIndex        =   11
            Top             =   1395
            Width           =   720
            _ExtentX        =   1270
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
         Begin MSMask.MaskEdBox Sequencial 
            Height          =   225
            Left            =   6705
            TabIndex        =   13
            Top             =   1350
            Width           =   945
            _ExtentX        =   1667
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
         Begin MSMask.MaskEdBox DataBaixaParc 
            Height          =   225
            Left            =   2625
            TabIndex        =   14
            Top             =   1380
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   5520
            TabIndex        =   15
            Top             =   600
            Width           =   1260
            _ExtentX        =   2223
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
         Begin VB.CheckBox ParcelaSelecionada 
            Height          =   255
            Left            =   4800
            TabIndex        =   49
            Top             =   360
            Width           =   735
         End
         Begin MSMask.MaskEdBox ValorTotalParcela 
            Height          =   225
            Left            =   5880
            TabIndex        =   57
            Top             =   960
            Width           =   1950
            _ExtentX        =   3440
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DescontoCancelar 
            Height          =   225
            Left            =   4800
            TabIndex        =   58
            Top             =   1200
            Width           =   1710
            _ExtentX        =   3016
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox JurosCancelar 
            Height          =   225
            Left            =   3090
            TabIndex        =   59
            Top             =   840
            Width           =   1410
            _ExtentX        =   2487
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MultaCancelar 
            Height          =   225
            Left            =   2460
            TabIndex        =   60
            Top             =   1170
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorCancelar 
            Height          =   225
            Left            =   360
            TabIndex        =   61
            Top             =   930
            Width           =   1740
            _ExtentX        =   3069
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1995
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3519
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox DataCreditoParc 
            Height          =   225
            Left            =   0
            TabIndex        =   198
            Top             =   0
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame FrameMovCCI 
         Caption         =   "Movimentos"
         Height          =   2415
         Left            =   480
         TabIndex        =   62
         Top             =   0
         Width           =   8415
         Begin VB.TextBox CtaCorrenteMovCCI 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   3000
            TabIndex        =   109
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox FilialEmpresaMov 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   960
            TabIndex        =   68
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton CancelarMov 
            Height          =   230
            Left            =   240
            TabIndex        =   67
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox HistoricoMovCCI 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   66
            Top             =   1320
            Width           =   5655
         End
         Begin VB.TextBox TipoMovCCI 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   4320
            TabIndex        =   63
            Top             =   960
            Width           =   1335
         End
         Begin MSMask.MaskEdBox ValorMovCCI 
            Height          =   225
            Left            =   5640
            TabIndex        =   64
            Top             =   960
            Width           =   1140
            _ExtentX        =   2011
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
         Begin MSMask.MaskEdBox DataMovCCI 
            Height          =   225
            Left            =   2040
            TabIndex        =   65
            Top             =   960
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMovimentosCCI 
            Height          =   2085
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3678
            _Version        =   393216
         End
      End
      Begin MSComctlLib.TabStrip TabDetalhes 
         Height          =   2895
         Left            =   510
         TabIndex        =   99
         Top             =   2520
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5106
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelamento"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhes"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   1
      Left            =   75
      TabIndex        =   6
      Top             =   720
      Width           =   9285
      Begin VB.Frame FrameFiltros 
         Caption         =   "Filtros"
         Height          =   5205
         Left            =   495
         TabIndex        =   7
         Top             =   60
         Width           =   8355
         Begin VB.Frame FrameFiltrosBaixaMovInt 
            Caption         =   "Filtros - Baixa a Baixa"
            Height          =   3975
            Left            =   360
            TabIndex        =   19
            Top             =   1000
            Width           =   7575
            Begin VB.Frame FrameFiltrosBaixaBaixa 
               BorderStyle     =   0  'None
               Caption         =   "Frame13"
               Height          =   1575
               Left            =   2640
               TabIndex        =   27
               Top             =   2280
               Width           =   4815
               Begin VB.Frame Frame6 
                  Caption         =   "Nº do Título"
                  Height          =   1575
                  Left            =   2520
                  TabIndex        =   35
                  Top             =   0
                  Width           =   2175
                  Begin MSMask.MaskEdBox TituloInic 
                     Height          =   300
                     Left            =   735
                     TabIndex        =   36
                     Top             =   435
                     Width           =   975
                     _ExtentX        =   1720
                     _ExtentY        =   529
                     _Version        =   393216
                     MaxLength       =   8
                     Mask            =   "99999999"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox TituloFim 
                     Height          =   300
                     Left            =   735
                     TabIndex        =   37
                     Top             =   960
                     Width           =   975
                     _ExtentX        =   1720
                     _ExtentY        =   529
                     _Version        =   393216
                     MaxLength       =   8
                     Mask            =   "99999999"
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
                     TabIndex        =   39
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
                     TabIndex        =   38
                     Top             =   990
                     Width           =   375
                  End
               End
               Begin VB.Frame Frame5 
                  Caption         =   "Data de Vencimento"
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   28
                  Top             =   0
                  Width           =   2175
                  Begin MSComCtl2.UpDown UpDownVencInic 
                     Height          =   300
                     Left            =   1725
                     TabIndex        =   29
                     TabStop         =   0   'False
                     Top             =   465
                     Width           =   240
                     _ExtentX        =   423
                     _ExtentY        =   529
                     _Version        =   393216
                     Enabled         =   -1  'True
                  End
                  Begin MSMask.MaskEdBox VencInic 
                     Height          =   300
                     Left            =   630
                     TabIndex        =   30
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
                  Begin MSComCtl2.UpDown UpDownVencFim 
                     Height          =   300
                     Left            =   1725
                     TabIndex        =   31
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
                     Left            =   630
                     TabIndex        =   32
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
                  Begin VB.Label Label1 
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
                     TabIndex        =   34
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
                     Left            =   195
                     TabIndex        =   33
                     Top             =   1020
                     Width           =   375
                  End
               End
            End
            Begin VB.Frame FrameTipoBaixas 
               Caption         =   "Tipo de Baixas"
               Height          =   975
               Left            =   360
               TabIndex        =   50
               Top             =   1200
               Width           =   6975
               Begin VB.OptionButton TipoBaixaCartao 
                  Caption         =   "Cartão"
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
                  TabIndex        =   185
                  Top             =   675
                  Width           =   975
               End
               Begin VB.OptionButton TipoBaixaPerdas 
                  Caption         =   "Perdas"
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
                  Left            =   5880
                  TabIndex        =   54
                  Top             =   345
                  Width           =   975
               End
               Begin VB.OptionButton TipoBaixaDebitos 
                  Caption         =   "Devoluções / Débitos"
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
                  Left            =   3600
                  TabIndex        =   53
                  Top             =   345
                  Width           =   2295
               End
               Begin VB.OptionButton TipoBaixaAdiantamentos 
                  Caption         =   "Adiantamentos"
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
                  Left            =   1920
                  TabIndex        =   52
                  Top             =   345
                  Width           =   1575
               End
               Begin VB.OptionButton TipoBaixaRecebimentos 
                  Caption         =   "Recebimentos"
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
                  TabIndex        =   51
                  Top             =   345
                  Value           =   -1  'True
                  Width           =   1575
               End
            End
            Begin VB.Frame FrameCliente 
               Caption         =   "Cliente"
               Height          =   885
               Left            =   360
               TabIndex        =   44
               Top             =   240
               Width           =   6975
               Begin VB.ComboBox Filial 
                  Height          =   315
                  Left            =   4650
                  TabIndex        =   45
                  Top             =   390
                  Width           =   1815
               End
               Begin MSMask.MaskEdBox Cliente 
                  Height          =   300
                  Left            =   1080
                  TabIndex        =   46
                  Top             =   390
                  Width           =   2400
                  _ExtentX        =   4233
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   20
                  PromptChar      =   "_"
               End
               Begin VB.Label LabelCliente 
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
                  Height          =   255
                  Left            =   210
                  MousePointer    =   14  'Arrow and Question
                  TabIndex        =   48
                  Top             =   435
                  Width           =   675
               End
               Begin VB.Label LabelFil 
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
                  Height          =   255
                  Left            =   4035
                  TabIndex        =   47
                  Top             =   435
                  Width           =   615
               End
            End
            Begin VB.Frame FrameDataBaixa 
               Caption         =   "Data da Baixa"
               Height          =   1575
               Left            =   360
               TabIndex        =   20
               Top             =   2280
               Width           =   2175
               Begin MSComCtl2.UpDown UpDownBaixaInic 
                  Height          =   300
                  Left            =   1710
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   435
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   -1  'True
               End
               Begin MSMask.MaskEdBox BaixaInic 
                  Height          =   300
                  Left            =   630
                  TabIndex        =   22
                  Top             =   450
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownBaixaFim 
                  Height          =   300
                  Left            =   1725
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   960
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   -1  'True
               End
               Begin MSMask.MaskEdBox BaixaFim 
                  Height          =   300
                  Left            =   630
                  TabIndex        =   24
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
               Begin VB.Label Label2 
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
                  Left            =   195
                  TabIndex        =   26
                  Top             =   990
                  Width           =   375
               End
               Begin VB.Label Label3 
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
                  TabIndex        =   25
                  Top             =   480
                  Width           =   375
               End
            End
            Begin VB.Frame FrameFiltrosMovInt 
               Caption         =   "Conta Corrente"
               Height          =   1575
               Left            =   2640
               TabIndex        =   40
               Top             =   1080
               Visible         =   0   'False
               Width           =   4695
               Begin VB.OptionButton CtaCorrenteApenas 
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
                  Left            =   360
                  TabIndex        =   43
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.OptionButton CtaCorrenteTodas 
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
                  Height          =   255
                  Left            =   360
                  TabIndex        =   42
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.ComboBox ContaCorrenteFiltro 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   41
                  Top             =   945
                  Width           =   2055
               End
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de cancelamento"
            Height          =   615
            Left            =   360
            TabIndex        =   16
            Top             =   240
            Width           =   7575
            Begin VB.OptionButton TipoCancMovimentoIntegral 
               Caption         =   "Movimento integral"
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
               Left            =   4440
               TabIndex        =   18
               Top             =   240
               Width           =   2415
            End
            Begin VB.OptionButton TipoCancBaixaBaixa 
               Caption         =   "Baixa a Baixa"
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
               Left            =   1200
               TabIndex        =   17
               Top             =   240
               Value           =   -1  'True
               Width           =   1575
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   3
      Left            =   75
      TabIndex        =   142
      Top             =   720
      Visible         =   0   'False
      Width           =   9285
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4440
         TabIndex        =   197
         Tag             =   "1"
         Top             =   2520
         Width           =   870
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3465
         TabIndex        =   155
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   154
         Top             =   2025
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   153
         Top             =   1635
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6375
         TabIndex        =   152
         Top             =   1590
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   147
         Top             =   3495
         Width           =   5895
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            TabIndex        =   151
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label CTBLabel7 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1125
            TabIndex        =   150
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   149
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   148
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
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
         Left            =   7875
         TabIndex        =   146
         Top             =   120
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6420
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   930
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
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
         Height          =   270
         Left            =   6420
         TabIndex        =   144
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
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
         Left            =   6420
         TabIndex        =   143
         Top             =   435
         Width           =   2700
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4665
         TabIndex        =   156
         Top             =   1335
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
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
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   157
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   158
         Top             =   1350
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   159
         Top             =   1290
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   160
         Top             =   1335
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
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
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   570
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   585
         TabIndex        =   162
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5580
         TabIndex        =   163
         Top             =   150
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3780
         TabIndex        =   164
         Top             =   120
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   0
         TabIndex        =   165
         Top             =   1185
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2985
         Left            =   6360
         TabIndex        =   166
         Top             =   1575
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6345
         TabIndex        =   167
         Top             =   1560
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel21 
         Caption         =   "Origem:"
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
         Left            =   45
         TabIndex        =   184
         Top             =   180
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   183
         Top             =   135
         Width           =   1530
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   4230
         TabIndex        =   182
         Top             =   615
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   181
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   180
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
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
         Left            =   1995
         TabIndex        =   179
         Top             =   600
         Width           =   870
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         TabIndex        =   178
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   177
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabelContas 
         Caption         =   "Plano de Contas"
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
         Left            =   6345
         TabIndex        =   176
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelTotais 
         Caption         =   "Totais:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   175
         Top             =   3060
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   174
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   173
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label CTBLabel8 
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
         Height          =   195
         Left            =   45
         TabIndex        =   172
         Top             =   570
         Width           =   480
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   171
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   170
         Top             =   180
         Width           =   450
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6360
         TabIndex        =   169
         Top             =   1290
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   6480
         TabIndex        =   168
         Top             =   750
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7125
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "BaixaRecCancelarOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   607
         Picture         =   "BaixaRecCancelarOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "BaixaRecCancelarOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5895
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10398
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Títulos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Baixas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
Attribute VB_Name = "BaixaRecCancelarOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTBaixaRecCancelar
Attribute objCT.VB_VarHelpID = -1

Private Sub CancelarMov_Click()
    Call objCT.CancelarMov_Click
End Sub

Private Sub ContaCorrenteFiltro_Click()
    Call objCT.ContaCorrenteFiltro_Click
End Sub

Private Sub DataCancelamento_Validate(Cancel As Boolean)
    Call objCT.DataCancelamento_Validate(Cancel)
End Sub

Private Sub LabelCliente_Click()
     Call objCT.LabelCliente_Click
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub BaixaInic_Change()
     Call objCT.BaixaInic_Change
End Sub

Private Sub BaixaInic_GotFocus()
     Call objCT.BaixaInic_GotFocus
End Sub

Private Sub BaixaInic_Validate(Cancel As Boolean)
     Call objCT.BaixaInic_Validate(Cancel)
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub TabDetalhes_Click()
    Call objCT.TabDetalhes_Click
End Sub

Private Sub TipoCancBaixaBaixa_Click()
    Call objCT.TipoCancBaixaBaixa_Click
End Sub

Private Sub TipoCancMovimentoIntegral_Click()
    Call objCT.TipoCancMovimentoIntegral_Click
End Sub

Private Sub UpDownBaixaInic_DownClick()
     Call objCT.UpDownBaixaInic_DownClick
End Sub

Private Sub UpDownBaixaInic_UpClick()
     Call objCT.UpDownBaixaInic_UpClick
End Sub

Private Sub BaixaFim_Change()
     Call objCT.BaixaFim_Change
End Sub

Private Sub BaixaFim_GotFocus()
     Call objCT.BaixaFim_GotFocus
End Sub

Private Sub BaixaFim_Validate(Cancel As Boolean)
     Call objCT.BaixaFim_Validate(Cancel)
End Sub

Private Sub UpDownBaixaFim_DownClick()
     Call objCT.UpDownBaixaFim_DownClick
End Sub

Private Sub UpDownBaixaFim_UpClick()
     Call objCT.UpDownBaixaFim_UpClick
End Sub

Private Sub UpDownVencInic_DownClick()
     Call objCT.UpDownVencInic_DownClick
End Sub

Private Sub UpDownVencInic_UpClick()
     Call objCT.UpDownVencInic_UpClick
End Sub

Private Sub UpDownVencFim_DownClick()
     Call objCT.UpDownVencFim_DownClick
End Sub

Private Sub UpDownVencFim_UpClick()
     Call objCT.UpDownVencFim_UpClick
End Sub

Private Sub TituloInic_Validate(Cancel As Boolean)
     Call objCT.TituloInic_Validate(Cancel)
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTBaixaRecCancelar
    Set objCT.objUserControl = Me
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

Private Sub VencFim_Change()
     Call objCT.VencFim_Change
End Sub

Private Sub VencFim_GotFocus()
     Call objCT.VencFim_GotFocus
End Sub

Private Sub VencFim_Validate(Cancel As Boolean)
     Call objCT.VencFim_Validate(Cancel)
End Sub

Private Sub TituloInic_Change()
     Call objCT.TituloInic_Change
End Sub

Private Sub TituloInic_GotFocus()
     Call objCT.TituloInic_GotFocus
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

Private Sub Recebimento_Click(Index As Integer)
     Call objCT.Recebimento_Click(Index)
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
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

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
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

Private Sub Numero_GotFocus()
     Call objCT.Numero_GotFocus
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
     Call objCT.Numero_KeyPress(KeyAscii)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
     Call objCT.Numero_Validate(Cancel)
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

Private Sub DataBaixaParc_GotFocus()
     Call objCT.DataBaixaParc_GotFocus
End Sub

Private Sub DataBaixaParc_KeyPress(KeyAscii As Integer)
     Call objCT.DataBaixaParc_KeyPress(KeyAscii)
End Sub

Private Sub DataBaixaParc_Validate(Cancel As Boolean)
     Call objCT.DataBaixaParc_Validate(Cancel)
End Sub

Private Sub ValorPagoBaixa_GotFocus()
     Call objCT.ValorPagoBaixa_GotFocus
End Sub

Private Sub ValorPagoBaixa_KeyPress(KeyAscii As Integer)
     Call objCT.ValorPagoBaixa_KeyPress(KeyAscii)
End Sub

Private Sub ValorPagoBaixa_Validate(Cancel As Boolean)
     Call objCT.ValorPagoBaixa_Validate(Cancel)
End Sub

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub Sequencial_GotFocus()
     Call objCT.Sequencial_GotFocus
End Sub

Private Sub Sequencial_KeyPress(KeyAscii As Integer)
     Call objCT.Sequencial_KeyPress(KeyAscii)
End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)
     Call objCT.Sequencial_Validate(Cancel)
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


Private Sub LabelFil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFil, Source, X, Y)
End Sub

Private Sub LabelFil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFil, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub ValorBaixado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorBaixado, Source, X, Y)
End Sub

Private Sub ValorBaixado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorBaixado, Button, Shift, X, Y)
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

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Desconto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Desconto, Source, X, Y)
End Sub

Private Sub Desconto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Desconto, Button, Shift, X, Y)
End Sub

Private Sub Multa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Multa, Source, X, Y)
End Sub

Private Sub Multa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Multa, Button, Shift, X, Y)
End Sub

Private Sub Juros_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Juros, Source, X, Y)
End Sub

Private Sub Juros_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Juros, Button, Shift, X, Y)
End Sub

Private Sub ValorPago_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPago, Source, X, Y)
End Sub

Private Sub ValorPago_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPago, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub DataBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataBaixa, Source, X, Y)
End Sub

Private Sub DataBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataBaixa, Button, Shift, X, Y)
End Sub

Private Sub SaldoDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoDebito, Source, X, Y)
End Sub

Private Sub SaldoDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoDebito, Button, Shift, X, Y)
End Sub

Private Sub NumTitulo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumTitulo, Source, X, Y)
End Sub

Private Sub NumTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumTitulo, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaCR_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaCR, Source, X, Y)
End Sub

Private Sub FilialEmpresaCR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaCR, Button, Shift, X, Y)
End Sub

Private Sub ValorDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorDebito, Source, X, Y)
End Sub

Private Sub ValorDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorDebito, Button, Shift, X, Y)
End Sub

Private Sub SiglaDocumentoCR_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SiglaDocumentoCR, Source, X, Y)
End Sub

Private Sub SiglaDocumentoCR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SiglaDocumentoCR, Button, Shift, X, Y)
End Sub

Private Sub DataEmissaoCred_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissaoCred, Source, X, Y)
End Sub

Private Sub DataEmissaoCred_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissaoCred, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub

Private Sub Label39_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label39, Source, X, Y)
End Sub

Private Sub Label39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label39, Button, Shift, X, Y)
End Sub

Private Sub Label38_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label38, Source, X, Y)
End Sub

Private Sub Label38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label38, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub CCIntNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CCIntNomeReduzido, Source, X, Y)
End Sub

Private Sub CCIntNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CCIntNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub ValorPA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPA, Source, X, Y)
End Sub

Private Sub ValorPA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPA, Button, Shift, X, Y)
End Sub

Private Sub MeioPagtoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MeioPagtoDescricao, Source, X, Y)
End Sub

Private Sub MeioPagtoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MeioPagtoDescricao, Button, Shift, X, Y)
End Sub

Private Sub DataMovimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataMovimento, Source, X, Y)
End Sub

Private Sub DataMovimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataMovimento, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Historico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Historico, Source, X, Y)
End Sub

Private Sub Historico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Historico, Button, Shift, X, Y)
End Sub

Private Sub Valor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Valor, Source, X, Y)
End Sub

Private Sub Valor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Valor, Button, Shift, X, Y)
End Sub

Private Sub ContaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaCorrente, Source, X, Y)
End Sub

Private Sub ContaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaCorrente, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub HistoricoPerda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(HistoricoPerda, Source, X, Y)
End Sub

Private Sub HistoricoPerda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(HistoricoPerda, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call objCT.Opcao_BeforeClick(Cancel)
End Sub

'Leo em 21/11/01 daqui p/ Baixo

Private Sub MultaCancelar_Change()
    Call objCT.MultaCancelar_Change
End Sub

Private Sub MultaCancelar_GotFocus()
    Call objCT.MultaCancelar_GotFocus
End Sub

Private Sub MultaCancelar_KeyPress(KeyAscii As Integer)
    Call objCT.MultaCancelar_KeyPress(KeyAscii)
End Sub

Private Sub MultaCancelar_Validate(Cancel As Boolean)
    Call objCT.MultaCancelar_Validate(Cancel)
End Sub

Private Sub JurosCancelar_Change()
    Call objCT.JurosCancelar_Change
End Sub

Private Sub JurosCancelar_GotFocus()
    Call objCT.JurosCancelar_GotFocus
End Sub

Private Sub JurosCancelar_KeyPress(KeyAscii As Integer)
    Call objCT.JurosCancelar_KeyPress(KeyAscii)
End Sub

Private Sub JurosCancelar_Validate(Cancel As Boolean)
    Call objCT.JurosCancelar_Validate(Cancel)
End Sub

Private Sub DescontoCancelar_Change()
    Call objCT.DescontoCancelar_Change
End Sub

Private Sub DescontoCancelar_GotFocus()
    Call objCT.DescontoCancelar_GotFocus
End Sub

Private Sub DescontoCancelar_KeyPress(KeyAscii As Integer)
    Call objCT.DescontoCancelar_KeyPress(KeyAscii)
End Sub

Private Sub DescontoCancelar_Validate(Cancel As Boolean)
    Call objCT.DescontoCancelar_Validate(Cancel)
End Sub

Private Sub ValorCancelar_Change()
    Call objCT.ValorCancelar_Change
End Sub

Private Sub ValorCancelar_GotFocus()
    Call objCT.ValorCancelar_GotFocus
End Sub

Private Sub ValorCancelar_KeyPress(KeyAscii As Integer)
    Call objCT.ValorCancelar_KeyPress(KeyAscii)
End Sub

Private Sub ValorCancelar_Validate(Cancel As Boolean)
    Call objCT.ValorCancelar_Validate(Cancel)
End Sub


'Private Sub FilialEmpresaBaixada_Change()
'    Call objCT.FilialEmpresaBaixada_Change
'End Sub
'
'Private Sub FilialEmpresaBaixada_GotFocus()
'    Call objCT.FilialEmpresaBaixada_GotFocus
'End Sub
'
'Private Sub FilialEmpresaBaixada_KeyPress(KeyAscii As Integer)
'    Call objCT.FilialEmpresaBaixada_KeyPress(KeyAscii)
'End Sub
'
'Private Sub FilialEmpresaBaixada_Validate(Cancel As Boolean)
'    Call objCT.FilialEmpresaBaixada_Validate(Cancel)
'End Sub

'Private Sub TipoBaixada_Change()
'    Call objCT.TipoBaixada_Change
'End Sub
'
'Private Sub TipoBaixada_GotFocus()
'    Call objCT.TipoBaixada_GotFocus
'End Sub
'
'Private Sub TipoBaixada_KeyPress(KeyAscii As Integer)
'    Call objCT.TipoBaixada_KeyPress(KeyAscii)
'End Sub
'
'Private Sub TipoBaixada_Validate(Cancel As Boolean)
'    Call objCT.TipoBaixada_Validate(Cancel)
'End Sub

'Private Sub NumeroBaixada_Change()
'    Call objCT.NumeroBaixada_Change
'End Sub
'
'Private Sub NumeroBaixada_GotFocus()
'    Call objCT.NumeroBaixada_GotFocus
'End Sub
'
'Private Sub NumeroBaixada_KeyPress(KeyAscii As Integer)
'    Call objCT.NumeroBaixada_KeyPress(KeyAscii)
'End Sub
'
'Private Sub NumeroBaixada_Validate(Cancel As Boolean)
'    Call objCT.NumeroBaixada_Validate(Cancel)
'End Sub

'Private Sub ParcelaBaixada_Change()
'    Call objCT.ParcelaBaixada_Change
'End Sub
'
'Private Sub ParcelaBaixada_GotFocus()
'    Call objCT.ParcelaBaixada_GotFocus
'End Sub
'
'Private Sub ParcelaBaixada_KeyPress(KeyAscii As Integer)
'    Call objCT.ParcelaBaixada_KeyPress(KeyAscii)
'End Sub
'
'Private Sub ParcelaBaixada_Validate(Cancel As Boolean)
'    Call objCT.ParcelaBaixada_Validate(Cancel)
'End Sub

'Private Sub ValorBaixada_Change()
'    Call objCT.ValorBaixada_Change
'End Sub
'
'Private Sub ValorBaixada_GotFocus()
'    Call objCT.ValorBaixada_GotFocus
'End Sub
'
'Private Sub ValorBaixada_KeyPress(KeyAscii As Integer)
'    Call objCT.ValorBaixada_KeyPress(KeyAscii)
'End Sub
'
'Private Sub ValorBaixada_Validate(Cancel As Boolean)
'    Call objCT.ValorBaixada_Validate(Cancel)
'End Sub

'Private Sub ValorParcelaBaixada_Change()
'    Call objCT.ValorParcelaBaixada_Change
'End Sub
'
'Private Sub ValorParcelaBaixada_GotFocus()
'    Call objCT.ValorParcelaBaixada_GotFocus
'End Sub
'
'Private Sub ValorParcelaBaixada_KeyPress(KeyAscii As Integer)
'    Call objCT.ValorParcelaBaixada_KeyPress(KeyAscii)
'End Sub
'
'Private Sub ValorParcelaBaixada_Validate(Cancel As Boolean)
'    Call objCT.ValorParcelaBaixada_Validate(Cancel)
'End Sub

'Private Sub SequencialBaixada_Change()
'    Call objCT.SequencialBaixada_Change
'End Sub
'
'Private Sub SequencialBaixada_GotFocus()
'    Call objCT.SequencialBaixada_GotFocus
'End Sub
'
'Private Sub SequencialBaixada_KeyPress(KeyAscii As Integer)
'    Call objCT.SequencialBaixada_KeyPress(KeyAscii)
'End Sub
'
'Private Sub SequencialBaixada_Validate(Cancel As Boolean)
'    Call objCT.SequencialBaixada_Validate(Cancel)
'End Sub

Private Sub MultaCancelar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MultaCancelar, Source, X, Y)
End Sub

Private Sub MultaCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MultaCancelar, Button, Shift, X, Y)
End Sub

Private Sub JurosCancelar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(JurosCancelar, Source, X, Y)
End Sub

Private Sub JurosCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(JurosCancelar, Button, Shift, X, Y)
End Sub

Private Sub DescontoCancelar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescontoCancelar, Source, X, Y)
End Sub

Private Sub DescontoCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescontoCancelar, Button, Shift, X, Y)
End Sub

Private Sub ValorCancelar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorCancelar, Source, X, Y)
End Sub

Private Sub ValorCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorCancelar, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaBaixada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaBaixada, Source, X, Y)
End Sub

Private Sub FilialEmpresaBaixada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaBaixada, Button, Shift, X, Y)
End Sub

Private Sub TipoBaixada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoBaixada, Source, X, Y)
End Sub

Private Sub TipoBaixada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoBaixada, Button, Shift, X, Y)
End Sub

Private Sub NumeroBaixada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroBaixada, Source, X, Y)
End Sub

Private Sub NumeroBaixada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroBaixada, Button, Shift, X, Y)
End Sub

Private Sub ParcelaBaixada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ParcelaBaixada, Source, X, Y)
End Sub

Private Sub ParcelaBaixada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ParcelaBaixada, Button, Shift, X, Y)
End Sub

Private Sub ValorParcelaBaixada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorParcelaBaixada, Source, X, Y)
End Sub

Private Sub ValorParcelaBaixada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorParcelaBaixada, Button, Shift, X, Y)
End Sub

Private Sub SequencialBaixada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SequencialBaixada, Source, X, Y)
End Sub

Private Sub SequencialBaixada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SequencialBaixada, Button, Shift, X, Y)
End Sub

Private Sub GridBaixasMovimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(GridBaixasMovimento, Source, X, Y)
End Sub

Private Sub GridBaixasMovimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(GridBaixasMovimento, Button, Shift, X, Y)
End Sub

Private Sub CancelarMov_Change()
    Call objCT.CancelarMov_Change
End Sub

Private Sub CancelarMov_GotFocus()
    Call objCT.CancelarMov_GotFocus
End Sub

Private Sub CancelarMov_KeyPress(KeyAscii As Integer)
    Call objCT.CancelarMov_KeyPress(KeyAscii)
End Sub

Private Sub CancelarMov_Validate(Cancel As Boolean)
    Call objCT.CancelarMov_Validate(Cancel)
End Sub

Private Sub CancelarMov_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CancelarMov, Source, X, Y)
End Sub

Private Sub CancelarMov_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CancelarMov, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaMov_GotFocus()
    Call objCT.FilialEmpresaMov_GotFocus
End Sub

Private Sub FilialEmpresaMov_KeyPress(KeyAscii As Integer)
    Call objCT.FilialEmpresaMov_KeyPress(KeyAscii)
End Sub

Private Sub FilialEmpresaMov_Validate(Cancel As Boolean)
    Call objCT.FilialEmpresaMov_Validate(Cancel)
End Sub

Private Sub FilialEmpresaMov_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaMov, Source, X, Y)
End Sub

Private Sub FilialEmpresaMov_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaMov, Button, Shift, X, Y)
End Sub

Private Sub HistoricoMovCCI_GotFocus()
    Call objCT.HistoricoMovCCI_GotFocus
End Sub

Private Sub HistoricoMovCCI_KeyPress(KeyAscii As Integer)
    Call objCT.HistoricoMovCCI_KeyPress(KeyAscii)
End Sub

Private Sub HistoricoMovCCI_Validate(Cancel As Boolean)
    Call objCT.HistoricoMovCCI_Validate(Cancel)
End Sub

Private Sub HistoricoMovCCI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(HistoricoMovCCI, Source, X, Y)
End Sub

Private Sub HistoricoMovCCI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(HistoricoMovCCI, Button, Shift, X, Y)
End Sub

Private Sub CtaCorrenteMovCCI_GotFocus()
    Call objCT.CtaCorrenteMovCCI_GotFocus
End Sub

Private Sub CtaCorrenteMovCCI_KeyPress(KeyAscii As Integer)
    Call objCT.CtaCorrenteMovCCI_KeyPress(KeyAscii)
End Sub

Private Sub CtaCorrenteMovCCI_Validate(Cancel As Boolean)
    Call objCT.HistoricoMovCCI_Validate(Cancel)
End Sub

Private Sub CtaCorrenteMovCCI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CtaCorrenteMovCCI, Source, X, Y)
End Sub

Private Sub CtaCorrenteMovCCI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CtaCorrenteMovCCI, Button, Shift, X, Y)
End Sub

Private Sub DataMovCCI_GotFocus()
    Call objCT.DataMovCCI_GotFocus
End Sub

Private Sub DataMovCCI_KeyPress(KeyAscii As Integer)
    Call objCT.DataMovCCI_KeyPress(KeyAscii)
End Sub

Private Sub DataMovCCI_Validate(Cancel As Boolean)
    Call objCT.DataMovCCI_Validate(Cancel)
End Sub

Private Sub DataMovCCI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataMovCCI, Source, X, Y)
End Sub

Private Sub DataMovCCI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataMovCCI, Button, Shift, X, Y)
End Sub

'Private Sub ValorMovCCI_Change()
'    Call objCT.ValorMovCCI_Change
'End Sub
'
'Private Sub ValorMovCCI_GotFocus()
'    Call objCT.ValorMovCCI_GotFocus
'End Sub
'
'Private Sub ValorMovCCI_KeyPress(KeyAscii As Integer)
'    Call objCT.ValorMovCCI_KeyPress(KeyAscii)
'End Sub
'
'Private Sub ValorMovCCI_Validate(Cancel As Boolean)
'    Call objCT.ValorMovCCI_Validate(Cancel)
'End Sub

Private Sub ValorMovCCI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorMovCCI, Source, X, Y)
End Sub

Private Sub ValorMovCCI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorMovCCI, Button, Shift, X, Y)
End Sub

'Private Sub TipoMovCCI_Change()
'    Call objCT.TipoMovCCI_Change
'End Sub
'
'Private Sub TipoMovCCI_GotFocus()
'    Call objCT.TipoMovCCI_GotFocus
'End Sub
'
'Private Sub TipoMovCCI_KeyPress(KeyAscii As Integer)
'    Call objCT.TipoMovCCI_KeyPress(KeyAscii)
'End Sub
'
'Private Sub TipoMovCCI_Validate(Cancel As Boolean)
'    Call objCT.TipoMovCCI_Validate(Cancel)
'End Sub

Private Sub TipoMovCCI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoMovCCI, Source, X, Y)
End Sub

Private Sub TipoMovCCI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoMovCCI, Button, Shift, X, Y)
End Sub

Private Sub GridMovimentosCCI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(GridMovimentosCCI, Source, X, Y)
End Sub

Private Sub GridMovimentosCCI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(GridMovimentosCCI, Button, Shift, X, Y)
End Sub


Private Sub GridMovimentosCCI_Click()
     Call objCT.GridMovimentosCCI_Click
End Sub

Private Sub GridMovimentosCCI_GotFocus()
     Call objCT.GridMovimentosCCI_GotFocus
End Sub

Private Sub GridMovimentosCCI_EnterCell()
     Call objCT.GridMovimentosCCI_EnterCell
End Sub

Private Sub GridMovimentosCCI_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridMovimentosCCI_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridMovimentosCCI_KeyPress(KeyAscii As Integer)
     Call objCT.GridMovimentosCCI_KeyPress(KeyAscii)
End Sub

Private Sub GridMovimentosCCI_Validate(Cancel As Boolean)
     Call objCT.GridMovimentosCCI_Validate(Cancel)
End Sub

Private Sub GridMovimentosCCI_RowColChange()
     Call objCT.GridMovimentosCCI_RowColChange
End Sub

Private Sub GridMovimentosCCI_Scroll()
     Call objCT.GridMovimentosCCI_Scroll
End Sub

Private Sub GridMovimentosCCI_LeaveCell()
     Call objCT.GridMovimentosCCI_LeaveCell
End Sub

Private Sub GridBaixasMovimento_Click()
     Call objCT.GridBaixasMovimento_Click
End Sub

Private Sub GridBaixasMovimento_GotFocus()
     Call objCT.GridBaixasMovimento_GotFocus
End Sub

Private Sub GridBaixasMovimento_EnterCell()
     Call objCT.GridBaixasMovimento_EnterCell
End Sub

Private Sub GridBaixasMovimento_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridBaixasMovimento_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridBaixasMovimento_KeyPress(KeyAscii As Integer)
     Call objCT.GridBaixasMovimento_KeyPress(KeyAscii)
End Sub

Private Sub GridBaixasMovimento_Validate(Cancel As Boolean)
     Call objCT.GridBaixasMovimento_Validate(Cancel)
End Sub

Private Sub GridBaixasMovimento_RowColChange()
     Call objCT.GridBaixasMovimento_RowColChange
End Sub

Private Sub GridBaixasMovimento_Scroll()
     Call objCT.GridBaixasMovimento_Scroll
End Sub

Private Sub GridBaixasMovimento_LeaveCell()
     Call objCT.GridBaixasMovimento_LeaveCell
End Sub

Private Sub ParcelaSelecionada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ParcelaSelecionada, Source, X, Y)
End Sub

Private Sub ParcelaSelecionada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ParcelaSelecionada, Button, Shift, X, Y)
End Sub

Private Sub ParcelaSelecionada_Click()
    Call objCT.ParcelaSelecionada_Click
End Sub

Private Sub ParcelaSelecionada_GotFocus()
    Call objCT.ParcelaSelecionada_GotFocus
End Sub

Private Sub ParcelaSelecionada_KeyPress(KeyAscii As Integer)
    Call objCT.ParcelaSelecionada_KeyPress(KeyAscii)
End Sub

Private Sub ParcelaSelecionada_Validate(Cancel As Boolean)
    Call objCT.ParcelaSelecionada_Validate(Cancel)
End Sub

Private Sub TipoCancMovimentoIntegral_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoCancMovimentoIntegral, Source, X, Y)
End Sub

Private Sub TipoCancMovimentoIntegral_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoCancMovimentoIntegral, Button, Shift, X, Y)
End Sub
Private Sub TipoCancBaixaBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoCancBaixaBaixa, Source, X, Y)
End Sub

Private Sub TipoCancBaixaBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoCancBaixaBaixa, Button, Shift, X, Y)
End Sub

Private Sub TipoBaixaRecebimentos_Click()
     Call objCT.TipoBaixaRecebimentos_Click
End Sub

Private Sub TipoBaixaAdiantamentos_Click()
     Call objCT.TipoBaixaAdiantamentos_Click
End Sub

Private Sub TipoBaixaDebitos_Click()
     Call objCT.TipoBaixaDebitos_Click
End Sub

Private Sub TipoBaixaPerdas_Click()
     Call objCT.TipoBaixaPerdas_Click
End Sub

Private Sub CtaCorrenteTodas_Click()
     Call objCT.CtaCorrenteTodas_Click
End Sub

Private Sub CtaCorrenteApenas_Click()
     Call objCT.CtaCorrenteApenas_Click
End Sub

Private Sub ContaCorrenteFiltro_Change()
    Call objCT.ContaCorrenteFiltro_Change
End Sub

Private Sub ContaCorrenteFiltro_Validate(Cancel As Boolean)
    Call objCT.ContaCorrenteFiltro_Validate(Cancel)
End Sub

Private Sub ItensSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ItensSelecionados, Source, X, Y)
End Sub

Private Sub ItensSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ItensSelecionados, Button, Shift, X, Y)
End Sub

Private Sub TotalCancelar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCancelar, Source, X, Y)
End Sub

Private Sub TotalCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCancelar, Button, Shift, X, Y)
End Sub

Private Sub ValorTotalParcela_GotFocus()
     Call objCT.ValorTotalParcela_GotFocus
End Sub

Private Sub ValorTotalParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorTotalParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorTotalParcela_Validate(Cancel As Boolean)
     Call objCT.ValorTotalParcela_Validate(Cancel)
End Sub

Private Sub ValorTotalParcela_Change()
    Call objCT.ValorTotalParcela_Change
End Sub

Private Sub DataCancelamento_Change()
    Call objCT.DataCancelamento_Change
End Sub

Private Sub HistoricoCancelamento_Change()
    Call objCT.HistoricoCancelamento_Change
End Sub

Private Sub UpDownDataCancelamento_UpClick()
    Call objCT.UpDownDataCancelamento_UpClick
End Sub

Private Sub UpDownDataCancelamento_DownClick()
    Call objCT.UpDownDataCancelamento_DownClick
End Sub

'contabilidade inicio

Private Sub CTBBotaoModeloPadrao_Click()
     Call objCT.CTBBotaoModeloPadrao_Click
End Sub

Private Sub CTBModelo_Click()
     Call objCT.CTBModelo_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub CTBGridContabil_EnterCell()
     Call objCT.CTBGridContabil_EnterCell
End Sub

Private Sub CTBGridContabil_GotFocus()
     Call objCT.CTBGridContabil_GotFocus
End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGridContabil_KeyPress(KeyAscii)
End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.CTBGridContabil_KeyDown(KeyCode, Shift)
End Sub

Private Sub CTBGridContabil_LeaveCell()
     Call objCT.CTBGridContabil_LeaveCell
End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)
     Call objCT.CTBGridContabil_Validate(Cancel)
End Sub

Private Sub CTBGridContabil_RowColChange()
     Call objCT.CTBGridContabil_RowColChange
End Sub

Private Sub CTBGridContabil_Scroll()
     Call objCT.CTBGridContabil_Scroll
End Sub

Private Sub CTBConta_Change()
     Call objCT.CTBConta_Change
End Sub

Private Sub CTBConta_GotFocus()
     Call objCT.CTBConta_GotFocus
End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)
     Call objCT.CTBConta_KeyPress(KeyAscii)
End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)
     Call objCT.CTBConta_Validate(Cancel)
End Sub

Private Sub CTBCcl_Change()
     Call objCT.CTBCcl_Change
End Sub

Private Sub CTBCcl_GotFocus()
     Call objCT.CTBCcl_GotFocus
End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCcl_KeyPress(KeyAscii)
End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)
     Call objCT.CTBCcl_Validate(Cancel)
End Sub

Private Sub CTBCredito_Change()
     Call objCT.CTBCredito_Change
End Sub

Private Sub CTBCredito_GotFocus()
     Call objCT.CTBCredito_GotFocus
End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCredito_KeyPress(KeyAscii)
End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)
     Call objCT.CTBCredito_Validate(Cancel)
End Sub

Private Sub CTBDebito_Change()
     Call objCT.CTBDebito_Change
End Sub

Private Sub CTBDebito_GotFocus()
     Call objCT.CTBDebito_GotFocus
End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBDebito_KeyPress(KeyAscii)
End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)
     Call objCT.CTBDebito_Validate(Cancel)
End Sub

Private Sub CTBSeqContraPartida_Change()
     Call objCT.CTBSeqContraPartida_Change
End Sub

Private Sub CTBSeqContraPartida_GotFocus()
     Call objCT.CTBSeqContraPartida_GotFocus
End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.CTBSeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.CTBSeqContraPartida_Validate(Cancel)
End Sub

Private Sub CTBHistorico_Change()
     Call objCT.CTBHistorico_Change
End Sub

Private Sub CTBHistorico_GotFocus()
     Call objCT.CTBHistorico_GotFocus
End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)
     Call objCT.CTBHistorico_KeyPress(KeyAscii)
End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)
     Call objCT.CTBHistorico_Validate(Cancel)
End Sub

Private Sub CTBLancAutomatico_Click()
    Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
    Call objCT.CTBAglutina_Click
End Sub

Private Sub CTBAglutina_GotFocus()
     Call objCT.CTBAglutina_GotFocus
End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)
     Call objCT.CTBAglutina_KeyPress(KeyAscii)
End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)
     Call objCT.CTBAglutina_Validate(Cancel)
End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_NodeClick(Node)
End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_Expand(Node)
End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwCcls_NodeClick(Node)
End Sub

Private Sub CTBListHistoricos_DblClick()
     Call objCT.CTBListHistoricos_DblClick
End Sub

Private Sub CTBBotaoLimparGrid_Click()
     Call objCT.CTBBotaoLimparGrid_Click
End Sub

Private Sub CTBLote_Change()
     Call objCT.CTBLote_Change
End Sub

Private Sub CTBLote_GotFocus()
     Call objCT.CTBLote_GotFocus
End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)
     Call objCT.CTBLote_Validate(Cancel)
End Sub

Private Sub CTBDataContabil_Change()
     Call objCT.CTBDataContabil_Change
End Sub

Private Sub CTBDataContabil_GotFocus()
     Call objCT.CTBDataContabil_GotFocus
End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)
     Call objCT.CTBDataContabil_Validate(Cancel)
End Sub

Private Sub CTBDocumento_Change()
     Call objCT.CTBDocumento_Change
End Sub

Private Sub CTBDocumento_GotFocus()
     Call objCT.CTBDocumento_GotFocus
End Sub

Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub CTBLabelDoc_Click()
     Call objCT.CTBLabelDoc_Click
End Sub

Private Sub CTBLabelLote_Click()
     Call objCT.CTBLabelLote_Click
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

'contabilidade fim

Private Sub TipoBaixaCartao_Click()
     Call objCT.TipoBaixaCartao_Click
End Sub


Private Sub CTBGerencial_Click()
    Call objCT.CTBGerencial_Click
End Sub

Private Sub CTBGerencial_GotFocus()
    Call objCT.CTBGerencial_GotFocus
End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)
    Call objCT.CTBGerencial_KeyPress(KeyAscii)
End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)
    Call objCT.CTBGerencial_Validate(Cancel)
End Sub

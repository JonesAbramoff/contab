VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RecebMaterialC 
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   ScaleHeight     =   5325
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4260
      Index           =   1
      Left            =   225
      TabIndex        =   0
      Top             =   825
      Width           =   9150
      Begin VB.Frame Frame11 
         Caption         =   "Entrada"
         Height          =   795
         Left            =   3300
         TabIndex        =   79
         Top             =   1185
         Width           =   5760
         Begin MSComCtl2.UpDown UpDownEntrada 
            Height          =   300
            Left            =   2700
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntrada 
            Height          =   300
            Left            =   1635
            TabIndex        =   5
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HoraEntrada 
            Height          =   300
            Left            =   4710
            TabIndex        =   6
            Top             =   315
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora Entrada:"
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
            Left            =   3435
            TabIndex        =   102
            Top             =   375
            Width           =   1200
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Data Entrada:"
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
            Left            =   360
            TabIndex        =   81
            Top             =   375
            Width           =   1200
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Nota Fiscal"
         Height          =   915
         Left            =   195
         TabIndex        =   74
         Top             =   2085
         Width           =   8865
         Begin VB.Frame FrameNFCliente 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   3930
            TabIndex        =   75
            Top             =   165
            Width           =   4410
            Begin VB.ComboBox Serie 
               Height          =   315
               Left            =   1035
               TabIndex        =   9
               Top             =   210
               Width           =   765
            End
            Begin MSMask.MaskEdBox NFiscal 
               Height          =   300
               Left            =   3225
               TabIndex        =   10
               Top             =   210
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin VB.Label SerieLabel 
               AutoSize        =   -1  'True
               Caption         =   "Série:"
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
               Left            =   525
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   77
               Top             =   270
               Width           =   510
            End
            Begin VB.Label NFiscalLabel 
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
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   2445
               TabIndex        =   76
               Top             =   240
               Width           =   720
            End
         End
         Begin VB.OptionButton NFiscalPropria 
            Caption         =   "Nota Fiscal Própria"
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
            Left            =   1155
            TabIndex        =   8
            Top             =   570
            Width           =   1980
         End
         Begin VB.OptionButton NFiscalCliente 
            Caption         =   "Nota Fiscal do Cliente"
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
            Left            =   1155
            TabIndex        =   7
            Top             =   270
            Width           =   2205
         End
         Begin VB.Frame FrameNFPropria 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   3930
            TabIndex        =   78
            Top             =   165
            Width           =   4410
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   795
         Left            =   195
         TabIndex        =   49
         Top             =   1185
         Width           =   2790
         Begin VB.CommandButton BotaoLimparRec 
            Height          =   315
            Left            =   1920
            Picture         =   "RecebMaterialC.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpar Código"
            Top             =   278
            Width           =   345
         End
         Begin VB.Label NumRecebimento 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1155
            TabIndex        =   73
            Top             =   285
            Width           =   765
         End
         Begin VB.Label LabelRecebimento 
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
            Left            =   420
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   72
            Top             =   330
            Width           =   660
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Nota Fiscal Original"
         Height          =   885
         Index           =   0
         Left            =   195
         TabIndex        =   46
         Top             =   150
         Width           =   8865
         Begin VB.ComboBox SerieNFiscalOriginal 
            Height          =   315
            Left            =   1755
            TabIndex        =   1
            Top             =   360
            Width           =   705
         End
         Begin VB.CommandButton BotaoExibirDados 
            Caption         =   "Exibir Dados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6480
            TabIndex        =   3
            Top             =   360
            Width           =   1485
         End
         Begin MSMask.MaskEdBox NFiscalOriginal 
            Height          =   315
            Left            =   4470
            TabIndex        =   2
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label SerieNFOriginalLabel 
            AutoSize        =   -1  'True
            Caption         =   "Série:"
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
            Left            =   1185
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   420
            Width           =   510
         End
         Begin VB.Label NFOriginalLabel 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3675
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   47
            Top             =   420
            Width           =   720
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dados do Cliente"
         Height          =   885
         Left            =   195
         TabIndex        =   43
         Top             =   3195
         Width           =   8865
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5460
            TabIndex        =   12
            Top             =   375
            Width           =   1860
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1680
            TabIndex        =   11
            Top             =   375
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4950
            TabIndex        =   45
            Top             =   420
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
            Height          =   165
            Left            =   990
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   44
            Top             =   420
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4275
      Index           =   3
      Left            =   240
      TabIndex        =   27
      Top             =   825
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame Frame4 
         Caption         =   "Dados de Transporte"
         Height          =   1215
         Left            =   120
         TabIndex        =   58
         Top             =   135
         Width           =   8910
         Begin VB.Frame Frame6 
            Caption         =   "Frete por conta"
            Height          =   795
            Index           =   1
            Left            =   465
            TabIndex        =   82
            Top             =   270
            Width           =   2220
            Begin VB.OptionButton Destinatario 
               Caption         =   "Destinatário"
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
               Left            =   405
               TabIndex        =   84
               Top             =   495
               Width           =   1695
            End
            Begin VB.OptionButton Emitente 
               Caption         =   "Emitente"
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
               Left            =   420
               TabIndex        =   83
               Top             =   225
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   7965
            TabIndex        =   30
            Top             =   735
            Width           =   735
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   5625
            MaxLength       =   10
            TabIndex        =   29
            Top             =   735
            Width           =   1290
         End
         Begin VB.ComboBox Transportadora 
            Height          =   315
            Left            =   4725
            TabIndex        =   28
            Top             =   270
            Width           =   2205
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "U.F. :"
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
            Left            =   7410
            TabIndex        =   61
            Top             =   795
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Placa Veículo:"
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
            Left            =   4275
            TabIndex        =   60
            Top             =   765
            Width           =   1275
         End
         Begin VB.Label TransportadoraLabel 
            AutoSize        =   -1  'True
            Caption         =   "Transportadora:"
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
            Left            =   3285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   59
            Top             =   330
            Width           =   1365
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Complemento"
         Height          =   1515
         Left            =   120
         TabIndex        =   54
         Top             =   2400
         Width           =   8910
         Begin VB.TextBox Mensagem 
            Height          =   300
            Left            =   2175
            MaxLength       =   250
            TabIndex        =   34
            Top             =   420
            Width           =   4710
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   300
            Left            =   5655
            TabIndex        =   36
            Top             =   1035
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   300
            Left            =   2160
            TabIndex        =   35
            Top             =   1035
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem N.Fiscal:"
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
            TabIndex        =   57
            Top             =   450
            Width           =   1725
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto:"
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
            Left            =   1095
            TabIndex        =   56
            Top             =   1088
            Width           =   1005
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Peso Líquido:"
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
            Left            =   4395
            TabIndex        =   55
            Top             =   1088
            Width           =   1200
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Volumes"
         Height          =   885
         Left            =   120
         TabIndex        =   50
         Top             =   1425
         Width           =   8910
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   5400
            TabIndex        =   33
            Top             =   338
            Width           =   1335
         End
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   3360
            TabIndex        =   32
            Top             =   338
            Width           =   1335
         End
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7275
            MaxLength       =   20
            TabIndex        =   85
            Top             =   345
            Width           =   1440
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1275
            TabIndex        =   31
            Top             =   345
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nº :"
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
            Left            =   6885
            TabIndex        =   86
            Top             =   398
            Width           =   345
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
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
            Left            =   135
            TabIndex        =   53
            Top             =   398
            Width           =   1050
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Espécie:"
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
            Left            =   2505
            TabIndex        =   52
            Top             =   398
            Width           =   750
         End
         Begin VB.Label Label32 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4800
            TabIndex        =   51
            Top             =   398
            Width           =   600
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   4275
      Index           =   2
      Left            =   225
      TabIndex        =   13
      Top             =   870
      Visible         =   0   'False
      Width           =   9150
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
         Height          =   345
         Left            =   90
         TabIndex        =   103
         Top             =   3750
         Width           =   1365
      End
      Begin VB.Frame Frame3 
         Caption         =   "Itens"
         Height          =   2685
         Left            =   45
         TabIndex        =   68
         Top             =   -15
         Width           =   9045
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   285
            Width           =   855
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6030
            MaxLength       =   50
            TabIndex        =   21
            Top             =   672
            Width           =   2595
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   225
            Left            =   4080
            TabIndex        =   18
            Top             =   1275
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   5070
            TabIndex        =   20
            Top             =   855
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
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   3810
            TabIndex        =   17
            Top             =   855
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorUnitario 
            Height          =   225
            Left            =   5100
            TabIndex        =   19
            Top             =   1290
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   228
            Left            =   2580
            TabIndex        =   16
            Top             =   288
            Width           =   996
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   180
            TabIndex        =   14
            Top             =   315
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotal 
            Height          =   225
            Left            =   6255
            TabIndex        =   22
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1860
            Left            =   105
            TabIndex        =   23
            Top             =   225
            Width           =   8820
            _ExtentX        =   15558
            _ExtentY        =   3281
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
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
         Left            =   5760
         TabIndex        =   25
         Top             =   3795
         Width           =   1365
      End
      Begin VB.Frame Frame9 
         Caption         =   "Valores"
         Height          =   930
         Left            =   60
         TabIndex        =   62
         Top             =   2700
         Width           =   9045
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   -20000
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   420
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   1569
            TabIndex        =   87
            Top             =   420
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   4527
            TabIndex        =   88
            Top             =   420
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   3048
            TabIndex        =   89
            Top             =   420
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IPIValor1 
            Height          =   285
            Left            =   90
            TabIndex        =   90
            Top             =   420
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
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
            Left            =   -20000
            TabIndex        =   70
            Top             =   210
            Width           =   825
         End
         Begin VB.Label SubTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6006
            TabIndex        =   92
            Top             =   420
            Width           =   1410
         End
         Begin VB.Label Total 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7485
            TabIndex        =   91
            Top             =   420
            Width           =   1410
         End
         Begin VB.Label LabelIPIValor 
            AutoSize        =   -1  'True
            Caption         =   "IPI"
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
            Left            =   668
            TabIndex        =   71
            Top             =   210
            Width           =   255
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Seguro"
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
            Left            =   3446
            TabIndex        =   67
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
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
            Left            =   2055
            TabIndex        =   66
            Top             =   210
            Width           =   450
         End
         Begin VB.Label LabelTotais 
            AutoSize        =   -1  'True
            Caption         =   "Valor Produtos"
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
            Left            =   6075
            TabIndex        =   65
            Top             =   210
            Width           =   1260
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Left            =   7965
            TabIndex        =   64
            Top             =   210
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Despesas"
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
            Left            =   4812
            TabIndex        =   63
            Top             =   210
            Width           =   840
         End
      End
      Begin VB.CommandButton BotaoCcls 
         Caption         =   "Centros de Custo"
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
         Left            =   7305
         TabIndex        =   26
         Top             =   3795
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4275
      Index           =   4
      Left            =   210
      TabIndex        =   93
      Top             =   855
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton BotaoLocalizacaoDist 
         Caption         =   "Estoque"
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
         Left            =   6960
         TabIndex        =   101
         Top             =   3885
         Width           =   1365
      End
      Begin VB.Frame Frame12 
         Caption         =   "Distribuição dos Produtos"
         Height          =   3660
         Left            =   300
         TabIndex        =   94
         Top             =   150
         Width           =   8595
         Begin VB.ComboBox ProdutoAlmoxDist 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   150
            Width           =   1920
         End
         Begin MSMask.MaskEdBox UMDist 
            Height          =   225
            Left            =   4425
            TabIndex        =   95
            Top             =   120
            Visible         =   0   'False
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxDist 
            Height          =   225
            Left            =   3060
            TabIndex        =   96
            Top             =   135
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantDist 
            Height          =   225
            Left            =   6540
            TabIndex        =   97
            Top             =   105
            Width           =   1470
            _ExtentX        =   2593
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
         Begin MSMask.MaskEdBox ItemNFDist 
            Height          =   225
            Left            =   525
            TabIndex        =   98
            Top             =   120
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDist 
            Height          =   2910
            Left            =   210
            TabIndex        =   99
            Top             =   345
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   5133
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox QuantItemNFDist 
            Height          =   225
            Left            =   5025
            TabIndex        =   100
            Top             =   150
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
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
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6765
      ScaleHeight     =   495
      ScaleWidth      =   2610
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   120
      Width           =   2670
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2100
         Picture         =   "RecebMaterialC.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RecebMaterialC.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RecebMaterialC.ctx":0BE2
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RecebMaterialC.ctx":0D6C
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   1605
         Picture         =   "RecebMaterialC.ctx":0EC6
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4695
      Left            =   180
      TabIndex        =   69
      Top             =   495
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribuição"
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
Attribute VB_Name = "RecebMaterialC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'************ GRADE **********************
Public gobjNFiscal As ClassNFiscal
'*****************************************

'Property Variables:
Dim m_Caption As String
Event Unload()

'distribuicao
Public iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFrameAtual As Integer

Dim objGrid As AdmGrid
'distribuicao
Public objGridItens As AdmGrid

'distribuicao
Public iGrid_Sequencial_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_Descricao_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
'distribuicao
'Dim iGrid_Almoxarifado_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_ValorTotal_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Ccl_Col As Integer

Dim WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Dim WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Dim WithEvents objEventoSerieNFOriginal As AdmEvento
Attribute objEventoSerieNFOriginal.VB_VarHelpID = -1
Dim WithEvents objEventoRecebimento As AdmEvento
Attribute objEventoRecebimento.VB_VarHelpID = -1
Dim WithEvents objEventoNFOriginal As AdmEvento
Attribute objEventoNFOriginal.VB_VarHelpID = -1
Dim WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Dim WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Dim WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
    
'distribuicao
Public gobjDistribuicao As Object
    
'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
'distribuicao
Private Const TAB_Distribuicao = 4

'distribuicao
Public Property Get GridItens() As Object
     Set GridItens = Me.Controls("GridItens")
End Property

Private Sub BotaoImprimir_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(NumRecebimento.Caption)) = 0 Then Error 57707
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Error 57708
      
    objNFiscal.sSerie = Serie.Text
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    objNFiscal.lNumRecebimento = StrParaLong(NumRecebimento.Caption)
    
    If NFiscalPropria.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRCP
    ElseIf NFiscalCliente.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRCC
    Else
        objNFiscal.iTipoNFiscal = 0
    End If
    
    objNFiscal.dtDataEntrada = StrParaDate(DataEntrada.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("NFiscalRec_Interna_TestaExistencia", objNFiscal)
    If lErro <> SUCESSO And lErro <> 57704 Then Error 57705
    If lErro = 57704 Then Error 57710
    
    lErro = objRelatorio.ExecutarDireto("Emissão das Notas de Recebimento", "Recebimento= @NRECEBIMENTO", 1, "NotaRec", "NRECEBIMENTO", objNFiscal.lNumRecebimento)
    If lErro <> SUCESSO Then Error 57709
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case Err
    
        Case 57705, 57709
        
        Case 57707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_PREENCHIDO", Err)
                
        Case 57708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", Err)
        
        Case 57710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFiscal.lNumNotaFiscal)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 166281)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimparRec_Click()

    NumRecebimento.Caption = ""
    
End Sub


Private Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoCcls_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_BotaoCcls_Click

    If GridItens.Row = 0 Then Error 52841

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then Error 52842
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub
    
Erro_BotaoCcls_Click:
    
    Select Case Err
    
        Case 52841
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 52842
             lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166282)

    End Select
    
    Exit Sub

End Sub

Private Sub DataEntrada_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEntrada, iAlterado)

End Sub

Private Sub Destinatario_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Emitente_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IPIValor1_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub IPIValor1_Validate(Cancel As Boolean)

    Call Valor_Saida(IPIValor1)

End Sub

Private Sub LabelRecebimento_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelRecebimento_Click

    'Se o Recebimento estiver preenchido
    If Len(Trim(NumRecebimento.Caption)) > 0 Then
        objNFiscal.lNumRecebimento = CLng(NumRecebimento.Caption)
    Else
        objNFiscal.lNumRecebimento = 0
    End If

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) <> 0 Then
        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 30624

        'Se não achou o Cliente --> erro
        If lErro = 12348 Then Error 30625

        objNFiscal.lCliente = objCliente.lCodigo

    End If

    If Len(Trim(Filial.Text)) <> 0 Then objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa

    If NFiscalPropria.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRCP

    ElseIf NFiscalCliente.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRCC

    Else
        objNFiscal.iTipoNFiscal = 0

    End If

    objNFiscal.sSerie = Serie.Text

    If Len(Trim(NFiscal.Text)) <> 0 Then objNFiscal.lNumNotaFiscal = CLng(NFiscal.Text)

    objNFiscal.dtDataEntrada = MaskedParaDate(DataEntrada)

    Call Chama_Tela("RecebMaterialCLista", colSelecao, objNFiscal, objEventoRecebimento)

    Exit Sub

Erro_LabelRecebimento_Click:

    Select Case Err

        Case 30624

        Case 30625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, Cliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 166283)

    End Select

    Exit Sub

End Sub

Private Sub NFiscal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)

End Sub

Private Sub NFiscal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscal_Validate

    If Len(Trim(NFiscal.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(NFiscal.Text)
        If lErro <> SUCESSO Then Error 57769
        
    End If
    
    Exit Sub
    
Erro_NFiscal_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 57769 'Erro tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166284)
          
    End Select

    Exit Sub


End Sub

Private Sub NFiscalOriginal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)

End Sub

Private Sub NFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalOriginal_Validate

    If Len(Trim(NFiscalOriginal.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(NFiscalOriginal.Text)
        If lErro <> SUCESSO Then Error 57762
        
    End If
    
    Exit Sub
    
Erro_NFiscalOriginal_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 57762 'Erro tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166285)
          
    End Select

    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) <> 0 And GridItens.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 52843

        'Coloca o valor do Ccl na coluna correspondente
        GridItens.TextMatrix(GridItens.Row, iGrid_Ccl_Col) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case Err

        Case 52843 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 166286)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
        
    If Len(Trim(NumRecebimento.Caption)) = 0 Then Error 61070
    
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Error 61122
    
    'Confirma exclusão?
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_RECEBIMENTO1", objNFiscal.lNumRecebimento)
    
    'Se resposta for sim
    If vbMsg = vbYes Then

        'Chama a rotina de Exclusão
        lErro = CF("RecebMaterialC_Exclui", objNFiscal)
        If lErro <> SUCESSO Then Error 42876

        'Limpa a tela
        Call Limpa_Tela_RecebMaterialC

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
        
        Case 42874, 42875, 42876, 61122

        Case 61070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166287)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExibirDados_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim bCancel As Boolean

On Error GoTo Erro_BotaoExibirDados_Click

    'Verifica se SerieNFiscalOriginal e NFiscalOriginal estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Or Len(Trim(NFiscalOriginal.Text)) = 0 Then Error 46625

    'Preenche a Serie e Número da Nota Fiscal
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)

    'Lê a NFiscal
    lErro = CF("NFiscalInternaSaidaNaoDev_Le_Numero", objNFiscal)
    If lErro <> SUCESSO And lErro <> 30765 Then Error 35644

    'Se não encontrou NFiscal Interna Saida Não Dev --> erro
    If lErro = 30765 Then Error 46887

    'Lê os Itens da NFiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then Error 46605

    
    Set gobjNFiscal = objNFiscal
    objNFiscal.lNumIntDoc = 0
    
    'Traz Dados comuns para Tela
    Cliente.Text = objNFiscal.lCliente

    Call Cliente_Validate(bCancel)

    Filial.Text = objNFiscal.iFilialCli

    Call Filial_Validate(bSGECancelDummy)

    Call Grid_Limpa(objGrid)

    lErro = Preenche_GridItens(objNFiscal.ColItensNF)
    If lErro <> SUCESSO Then Error 46604


    lErro = SubTotal_Calcula()
    If lErro <> SUCESSO Then Error 46577

    Exit Sub

Erro_BotaoExibirDados_Click:

    Select Case Err

        Case 46625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NUMERO_ORIGINAL_FALTANDO", Err)

        Case 35644, 46605, 46604, 46577

        Case 46887
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_INT_SAIDA_NAO_DEV_NAO_CADASTRADA", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166288)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 41788

    'Limpa a tela
    Call Limpa_Tela_RecebMaterialC

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 41788

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166289)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 33643

    'Limpa a Tela
    Call Limpa_Tela_RecebMaterialC

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 33643

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166290)

    End Select

    Exit Sub

End Sub


Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As New Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 52191

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 33645
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto
    
    Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 33645

        Case 52191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166291)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    iClienteAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        'Verifica se Cliente foi preenchido
        If Len(Trim(Cliente.Text)) > 0 Then

            'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then Error 33646

            'Lê coleção de códigos, nomes de Filiais do Cliente
            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then Error 30647

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)

               'Se Cliente não foi preenchido --> Limpa combo de Filiais
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            Filial.Clear

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case Err

        Case 30647
        
        Case 33646
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166292)

    End Select

    Exit Sub

End Sub

Private Sub DataEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrada_Validate

    'Verifica se Data Entrada foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Critica a Data
    lErro = Data_Critica(DataEntrada.Text)
    If lErro <> SUCESSO Then Error 30509

    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True


    Select Case Err

        Case 30509

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166293)

    End Select

    Exit Sub

End Sub

'horaentrada
Public Sub HoraEntrada_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HoraEntrada, iAlterado)

End Sub

'horaentrada
Public Sub HoraEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'horaentrada
Public Sub HoraEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraEntrada_Validate

    'Verifica se a hora de Entrada foi digitada
    If Len(Trim(HoraEntrada.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraEntrada.Text)
    If lErro <> SUCESSO Then gError 89809

    Exit Sub

Erro_HoraEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 89809

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166294)

    End Select

    Exit Sub

End Sub

Private Sub DescricaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 30648

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then Error 30649

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 30650

        'Se não encontrou a Filial do Cliente --> erro
        If lErro = 17660 Then Error 30651

        'coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 30652

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 30648, 30650

        Case 30651
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 30649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 30652
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166295)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim vCodigo As Variant
Dim colSerie As New colSerie
Dim colSiglasUF As New Collection
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objEventoCliente = New AdmEvento
    Set objEventoSerie = New AdmEvento
    Set objEventoSerieNFOriginal = New AdmEvento
    Set objEventoRecebimento = New AdmEvento
    Set objEventoNFOriginal = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoCcl = New AdmEvento

    'distribuicao
    Set gobjDistribuicao = CreateObject("RotinasMat.ClassMATDist")
    Set gobjDistribuicao.objTela = Me
    gobjDistribuicao.bTela = True


    'Lê as séries correspondentes a FilialEmpresa = giFilialEmpresa
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 30603

    'Preenche a List das Combos Serie e SerieNFiscalOriginal
    For iIndice = 1 To colSerie.Count

        SerieNFiscalOriginal.AddItem colSerie(iIndice).sSerie
        Serie.AddItem colSerie(iIndice).sSerie

    Next

    'Lê as siglas dos Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colSiglasUF, STRING_ESTADO_SIGLA)
    If lErro <> SUCESSO Then gError 30604

    'Alimenta a Combo PlacaUF.
    For iIndice = 1 To colSiglasUF.Count

        PlacaUF.AddItem colSiglasUF(iIndice)

    Next

    'Lê o código e o Nome Reduzido da Transportadora
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 30605

    'Preenche a Combo Box Transportadora com código e Nome Reduzido
    For iIndice = 1 To colCodigoDescricao.Count

        Transportadora.AddItem colCodigoDescricao(iIndice).iCodigo & "-" & colCodigoDescricao(iIndice).sNome

        'Preenche ItemData com o Código
        Transportadora.ItemData(Transportadora.NewIndex) = colCodigoDescricao(iIndice).iCodigo

    Next

    'Incluído por Luiz Nogueira em 21/08/03
    'Carrega a combo VolumeEspecie
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
    If lErro <> SUCESSO Then gError 102430

    'Incluído por Luiz Nogueira em 21/08/03
    'Carrega a combo VolumeMarca
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    If lErro <> SUCESSO Then gError 102431

    'Coloca gdtDataAtual em DataEntrada
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True

    'Inicializa a Mascára de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 30606

    'Inicializa a Mascara de Ccl
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then gError 49379

    'Formata a Quantidade para o Formato de Estoque
    Quantidade.Format = FORMATO_ESTOQUE

    Set objGrid = New AdmGrid
'distribuicao
    Set objGridItens = objGrid

    'Inicializa GridItens
    lErro = Inicializa_GridItens(objGrid)
    If lErro <> SUCESSO Then gError 30607
            
'distribuicao
    'Inicializa o grid de Distribuicao
    lErro = gobjDistribuicao.Inicializa_GridDist()
    If lErro <> SUCESSO Then gError 89597

    NFiscalPropria.Value = True
        
    iAlterado = 0

    Set gobjNFiscal = New ClassNFiscal
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 30603, 30604, 30605, 30606, 30607, 49379, 89597, 102430, 102431

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166296)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_MascaraCcl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 49380

    Ccl.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_MascaraCcl:

    Inicializa_MascaraCcl = Err

    Select Case Err

        Case 49380

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166297)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim iIndice As Integer
Dim vCodigo As Variant
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RecebCli"

    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Error 30608

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_SERIE, "Serie"
    colCampoValor.Add "NumRecebimento", objNFiscal.lNumRecebimento, 0, "NumRecebimento"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "ValorProdutos", objNFiscal.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorFrete", objNFiscal.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorSeguro", objNFiscal.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorOutrasDespesas", objNFiscal.dValorOutrasDespesas, 0, "ValorOutrasDespesas"
    colCampoValor.Add "ValorDesconto", objNFiscal.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorTotal", objNFiscal.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_NFISCAL_PLACA, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_NFISCAL_PLACA_UF, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, 0, "VolumeEspecie" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, 0, "VolumeMarca" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_NFISCAL_MENSAGEM, "MensagemNota"
    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "PesoLiq"
    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "DataEntrada", objNFiscal.dtDataEntrada, 0, "DataEntrada"
'horaentrada
    colCampoValor.Add "HoraEntrada", CDbl(objNFiscal.dtHoraEntrada), 0, "HoraEntrada"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_IGUAL, STATUS_LANCADO

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 30608

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166298)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objNFiscal.dtDataEntrada = colCampoValor.Item("DataEntrada").vValor
'horaentrada
    objNFiscal.dtHoraEntrada = colCampoValor.Item("HoraEntrada").vValor
    objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objNFiscal.lCliente = colCampoValor.Item("Cliente").vValor
    objNFiscal.iFilialCli = colCampoValor.Item("FilialCli").vValor
    objNFiscal.iTipoNFiscal = colCampoValor.Item("TipoNFiscal").vValor
    objNFiscal.sSerie = colCampoValor.Item("Serie").vValor
    objNFiscal.lNumRecebimento = colCampoValor.Item("NumRecebimento").vValor
    objNFiscal.lNumNotaFiscal = colCampoValor.Item("NumNotaFiscal").vValor
    objNFiscal.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
    objNFiscal.dValorFrete = colCampoValor.Item("ValorFrete").vValor
    objNFiscal.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
    objNFiscal.dValorOutrasDespesas = colCampoValor.Item("ValorOutrasDespesas").vValor
    objNFiscal.dValorDesconto = colCampoValor.Item("ValorDesconto").vValor
    objNFiscal.dValorTotal = colCampoValor.Item("ValorTotal").vValor
    objNFiscal.iCodTransportadora = colCampoValor.Item("CodTransportadora").vValor
    objNFiscal.sPlaca = colCampoValor.Item("Placa").vValor
    objNFiscal.sPlacaUF = colCampoValor.Item("PlacaUF").vValor
    objNFiscal.lVolumeQuant = colCampoValor.Item("VolumeQuant").vValor
    objNFiscal.lVolumeEspecie = colCampoValor.Item("VolumeEspecie").vValor
    objNFiscal.lVolumeMarca = colCampoValor.Item("VolumeMarca").vValor
    objNFiscal.sMensagemNota = colCampoValor.Item("MensagemNota").vValor
    objNFiscal.dPesoLiq = colCampoValor.Item("PesoLiq").vValor
    objNFiscal.dPesoBruto = colCampoValor.Item("PesoBruto").vValor
    objNFiscal.iFreteRespons = colCampoValor.Item("FreteRespons").vValor
    objNFiscal.sVolumeNumero = colCampoValor.Item("VolumeNumero").vValor

    'Lê NFiscal no BD
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 89240

    If lErro = 31442 Then gError 89241

    'Chama Preenche_Tela
    lErro = Preenche_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 30609

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 30609, 89240

        Case 89241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_NAO_CADASTRADO", gErr, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166299)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objNFiscal As ClassNFiscal, Optional iGravacao = 1) As Long
'Move os dados da tela para memória

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_Move_Tela_Memoria

    'Status
    objNFiscal.iStatus = STATUS_LANCADO
    
    'Verifica se o Recebimento está preenchido
    If Len(Trim(NumRecebimento.Caption)) > 0 Then
        'Se estiver coloca no obj
        objNFiscal.lNumRecebimento = CLng(NumRecebimento.Caption)
    Else
        objNFiscal.lNumRecebimento = 0
    End If
    
    'Lê Cliente no BD
    If Len(Trim(Cliente.Text)) <> 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 30610

        'Se não achou o Cliente --> erro
        If lErro = 12348 Then Error 30611

        objNFiscal.lCliente = objCliente.lCodigo

    End If

    If Len(Trim(Filial.Text)) <> 0 Then objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)

    objNFiscal.dtDataEntrada = MaskedParaDate(DataEntrada)

'horaentrada
    If Len(Trim(HoraEntrada.ClipText)) > 0 Then
        objNFiscal.dtHoraEntrada = CDate(HoraEntrada.Text)
    Else
        objNFiscal.dtHoraEntrada = Time
    End If

    If NFiscalPropria.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRCP
        objNFiscal.sSerie = ""
        objNFiscal.lNumNotaFiscal = 0

    ElseIf NFiscalCliente.Value Then
        objNFiscal.iTipoNFiscal = DOCINFO_NRCC
        objNFiscal.sSerie = Serie.Text
        If Len(Trim(NFiscal.Text)) <> 0 Then objNFiscal.lNumNotaFiscal = CLng(NFiscal.Text)
    Else
        objNFiscal.iTipoNFiscal = 0
    End If

    objNFiscal.iFilialEmpresa = giFilialEmpresa
    If Len(Trim(SubTotal.Caption)) <> 0 Then objNFiscal.dValorProdutos = CDbl(SubTotal.Caption)
    If Len(Trim(Total.Caption)) <> 0 Then objNFiscal.dValorTotal = CDbl(Total.Caption)
    If Len(Trim(ValorDesconto.Text)) <> 0 Then objNFiscal.dValorDesconto = CDbl(ValorDesconto.Text)
    If Len(Trim(ValorSeguro.Text)) <> 0 Then objNFiscal.dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(ValorFrete.Text)) <> 0 Then objNFiscal.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorDespesas.Text)) <> 0 Then objNFiscal.dValorOutrasDespesas = CDbl(ValorDespesas.Text)

    'Armazena o responsável pelo frete
    If Emitente.Value Then
        objNFiscal.iFreteRespons = FRETE_EMITENTE
    Else
        objNFiscal.iFreteRespons = FRETE_DESTINATARIO
    End If
    
    'Preenche a coleção de Ítens
    lErro = Move_Grid_Memoria(objNFiscal.ColItensNF)
    If lErro <> SUCESSO Then Error 30612

    objNFiscal.lNumIntDoc = 0

    objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text

    If Len(Trim(VolumeQuant.Text)) <> 0 Then objNFiscal.lVolumeQuant = CInt(VolumeQuant.Text)

    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.sMensagemNota = Mensagem.Text
    
    If Len(Trim(PesoLiquido.Text)) <> 0 Then objNFiscal.dPesoLiq = CDbl(PesoLiquido.Text)
    If Len(Trim(PesoBruto.Text)) <> 0 Then objNFiscal.dPesoBruto = CDbl(PesoBruto.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 30610, 30612

        Case 30611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166300)

    End Select

    Exit Function

End Function

Private Function Move_Grid_Memoria(colItensNFiscal As ColItensNF) As Long
'Move os Itens do Grid para a Memória

Dim iIndice As Integer
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim colAlocacoes As ColAlocacoesItemNF
Dim objProduto As New ClassProduto
Dim dValorDesconto As Double
Dim dPercentDesc As Double
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Move_Grid_Memoria

    For iIndice = 1 To objGrid.iLinhasExistentes

        sProdutoFormatado = ""

        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30613

        objProduto.sCodigo = sProdutoFormatado

        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 25178

        'Se não achou o Produto --> erro
        If lErro = 28030 Then gError 25179

        'Verifica se Ccl foi preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Ccl_Col))) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", GridItens.TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 49409

        Else
            sCclFormatada = ""
        End If

        dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        dPercentDesc = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col))

        'distribuicao. retirado o codigo do almoxarifado
        colItensNFiscal.Add 0, iIndice, sProdutoFormatado, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)), StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)), dPercentDesc, dValorDesconto, DATA_NULA, GridItens.TextMatrix(iIndice, iGrid_Descricao_Col), 0, 0, 0, 0, 0, colAlocacoes, 0, "", sCclFormatada, 0, 0, "", 0, 0, objProduto.iControleEstoque, objProduto.sSiglaUMEstoque, objProduto.iClasseUM, 0
        
        '********************* TRATAMENTO DE GRADE *****************
        lErro = gobjDistribuicao.Move_DistribuicaoGrade_Memoria(gobjNFiscal.ColItensNF(iIndice))
        If lErro <> SUCESSO Then gError 86375
        
        Call Move_ItensGrade_Tela(colItensNFiscal(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
        
    Next

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = gErr

    Select Case gErr

        Case 25178, 30613, 49409, 86375

        Case 25179
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166301)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    
'distribuicao
'    objGridInt.colColuna.Add ("Almoxarifado")
    
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Valor Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Valor Total")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    
'distribuicao
'    objGridInt.colCampo.Add (Almoxarifado.Name)

    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ValorUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)

    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
'distribuicao
'    iGrid_Almoxarifado_Col = 5
    iGrid_Ccl_Col = 5
    iGrid_ValorUnitario_Col = 6
    iGrid_PercDesc_Col = 7
    iGrid_Desconto_Col = 8
    iGrid_ValorTotal_Col = 9
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_RECEB + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Private Function Preenche_GridItens(colItens As ColItensNF) As Long
'Preenche o Grid de Itens com os objetos da coleção colItens

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dValorTotal As Double
Dim objItemNFiscal As ClassItemNF
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sCclMascarado As String

On Error GoTo Erro_Preenche_GridItens

    objGrid.iLinhasExistentes = colItens.Count

    'Preenche GridItens
    For Each objItemNFiscal In colItens

        iIndice = iIndice + 1

        sProdutoEnxuto = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoEnxuto(objItemNFiscal.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 30616

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objItemNFiscal.colItensRomaneioGrade.Count > 0 Then GridItens.TextMatrix(iIndice, 0) = "# " & GridItens.TextMatrix(iIndice, 0)
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemNFiscal.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNFiscal.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNFiscal.dQuantidade)
        If objItemNFiscal.dPercDesc <> 0 Then GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNFiscal.dPercDesc, "Percent")
        If objItemNFiscal.dValorDesconto <> 0 Then GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNFiscal.dValorDesconto, "Standard")
             
        sCclMascarado = ""
        
        'mascara Ccl , se estiver informada
        If objItemNFiscal.sCcl <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_RetornaCclEnxuta(objItemNFiscal.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then Error 49410
            
            'Preenche o campo Ccl com o Ccl encontrado
            Ccl.PromptInclude = False
            Ccl.Text = sCclMascarado
            Ccl.PromptInclude = True
    
            'Joga o Ccl encontrado no Grid
            GridItens.TextMatrix(iIndice, iGrid_Ccl_Col) = Ccl.Text

        End If
            
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(CStr(objItemNFiscal.dPrecoUnitario), "Standard")

        'Calcula o Valor Total a partir do Valor Unitário e Quantidade
        dValorTotal = (objItemNFiscal.dPrecoUnitario * objItemNFiscal.dQuantidade) - objItemNFiscal.dValorDesconto

        'Coloca o Valor Total na Coluna correspondente
        GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col) = Format(CStr(dValorTotal), "Standard")

''''        Call gobjDistribuicao.Atualiza
    Next

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = Err

    Select Case Err

        Case 29038, 30616, 49410

        Case 29039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166302)

    End Select

    Exit Function

End Function

Function Preenche_Tela(objNFiscal As ClassNFiscal) As Long
'Preenche os campos da tela com os dados passados como parâmetro em objNFiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objNFiscalOriginal As New ClassNFiscal

On Error GoTo Erro_Preenche_Tela

    'Lê os ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 30617

    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 89604

    If objNFiscal.lNumIntNotaOriginal > 0 Then
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal

        'Lê a NFiscal
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 30618

        'Se não achou a NFiscal --> erro
        If lErro = 31442 Then gError 30623

    End If

    Call Limpa_Tela_RecebMaterialC

    lErro = Coloca_Dados_Tela(objNFiscal, objNFiscalOriginal)
    If lErro <> SUCESSO Then gError 30525
    
    Set gobjNFiscal = objNFiscal
    
    iAlterado = 0

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 30617, 30618, 30525, 89604

        Case 30623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA", gErr, objNFiscalOriginal.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166303)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set gobjNFiscal = Nothing
    Set objEventoNFOriginal = Nothing
    Set objEventoSerieNFOriginal = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoCliente = Nothing
    Set objEventoRecebimento = Nothing
    Set objEventoSerie = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCcl = Nothing

    Set objGrid = Nothing

    'distribuicao
    Set gobjDistribuicao = Nothing

    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub BotaoLocalizacaoDist_Click()
'distribuicao

    Call gobjDistribuicao.BotaoLocalizacaoDist_Click

End Sub

Public Sub ItemNFDist_Change()
'distribuicao

    Call gobjDistribuicao.ItemNFDist_Change

End Sub

Public Sub ItemNFDist_GotFocus()
'distribuicao

    Call gobjDistribuicao.ItemNFDist_GotFocus

End Sub

Public Sub ItemNFDist_KeyPress(KeyAscii As Integer)
'distribuicao

    Call gobjDistribuicao.ItemNFDist_KeyPress(KeyAscii)

End Sub

Public Sub ItemNFDist_Validate(Cancel As Boolean)
'distribuicao

    Call gobjDistribuicao.ItemNFDist_Validate(Cancel)

End Sub

Public Sub AlmoxDist_Change()
'distribuicao

    Call gobjDistribuicao.AlmoxDist_Change

End Sub

Public Sub AlmoxDist_GotFocus()
'distribuicao

    Call gobjDistribuicao.AlmoxDist_GotFocus

End Sub

Public Sub AlmoxDist_KeyPress(KeyAscii As Integer)
'distribuicao

    Call gobjDistribuicao.AlmoxDist_KeyPress(KeyAscii)

End Sub

Public Sub AlmoxDist_Validate(Cancel As Boolean)
'distribuicao

    Call gobjDistribuicao.AlmoxDist_Validate(Cancel)

End Sub

Public Sub QuantDist_Change()
'distribuicao

    Call gobjDistribuicao.QuantDist_Change

End Sub

Public Sub QuantDist_GotFocus()
'distribuicao

    Call gobjDistribuicao.QuantDist_GotFocus

End Sub

Public Sub QuantDist_KeyPress(KeyAscii As Integer)
'distribuicao

    Call gobjDistribuicao.QuantDist_KeyPress(KeyAscii)

End Sub

Public Sub QuantDist_Validate(Cancel As Boolean)
'distribuicao

    Call gobjDistribuicao.QuantDist_Validate(Cancel)

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub Mensagem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NFiscalCliente_Click()

    iAlterado = REGISTRO_ALTERADO
    
    FrameNFCliente.Visible = True
    FrameNFPropria.Visible = False
    
End Sub

Private Sub NFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NFiscalPropria_Click()

    iAlterado = REGISTRO_ALTERADO
    
    FrameNFPropria.Visible = True
    FrameNFCliente.Visible = False
        
End Sub

Private Sub NFOriginalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection

    objNFiscal.sSerie = SerieNFiscalOriginal.Text

    If Len(Trim(NFiscalOriginal.Text)) <> 0 Then objNFiscal.lNumIntDoc = CLng(NFiscalOriginal.Text)

    'Passar em colSelecao o Destinatario
    colSelecao.Add DOCINFO_CLIENTE

    'Trocar para NFInternaDestinatarioLista
    Call Chama_Tela("NFInternaDestinatarioLista", colSelecao, objNFiscal, objEventoNFOriginal)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(bCancel)

    Me.Show

End Sub


Private Sub objEventoRecebimento_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoRecebimento_evSelecao

    Set objNFiscal = obj1
    
    'Lê NFiscal no BD
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 89244
    
    If lErro = 31442 Then gError 89245
    
    'Chama Preenche_Tela
    lErro = Preenche_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 30628

    'Fecha o os comandos se seta se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoRecebimento_evSelecao:

    Select Case gErr

        Case 30628, 89244

        Case 89245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_NAO_CADASTRADO", gErr, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166304)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNFOriginal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

    Set objNFiscal = obj1

    SerieNFiscalOriginal.Text = objNFiscal.sSerie

    NFiscalOriginal.Text = CStr(objNFiscal.lNumNotaFiscal)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    If GridItens.Row = 0 Then Exit Sub

    Set objProduto = obj1

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 30631

    'Verifica se a Coluna do Produto é a corrente e se o Produto não está preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

        sProdutoEnxuto = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 30632

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 30633

        If lErro = 28030 Then Error 30634

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        If Not (Me.ActiveControl Is Produto) Then

            'Preenche a célula de Produto
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then Error 30635

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case Err

        Case 30631, 30632, 30633, 30635

        Case 30634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166305)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    Serie.Text = objSerie.sSerie

    Me.Show

End Sub

Private Sub objEventoSerieNFOriginal_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    SerieNFiscalOriginal.Text = objSerie.sSerie

    Me.Show

End Sub

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    Transportadora.Text = objTransportadora.iCodigo & "-" & objTransportadora.sNomeReduzido

    Me.Show

End Sub

Private Sub Placa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PlacaUF_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PlacaUF_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Serie_Validate

    'Verifica se a Série está preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub

    'Verifica se é uma Série selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub

    'Verifica se NFiscalPropria
    If NFiscalPropria.Value = True Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(Serie)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 30653

        If lErro = 12253 Then Error 30654

    Else
    
        'Verifica se tamanho da série é maior do que o espaço no bd ==> erro
        If Len(Trim(Serie.Text)) > STRING_SERIE Then Error 30655

    End If

    Exit Sub

Erro_Serie_Validate:

    Cancel = True


    Select Case Err

        Case 30653

        Case 30654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)

        Case 30655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166306)

    End Select

    Exit Sub

End Sub

Private Sub SerieLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    If Len(Trim(Serie.Text)) <> 0 Then objSerie.sSerie = Serie.Text

    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

End Sub

Private Sub SerieNFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SerieNFiscalOriginal_Click()

Dim lNumAuto As Long
Dim lErro As Long

On Error GoTo Erro_SerieNFiscalOriginal_Click

    iAlterado = REGISTRO_ALTERADO

    If SerieNFiscalOriginal.ListIndex = -1 Then Exit Sub

    Exit Sub

Erro_SerieNFiscalOriginal_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166307)

    End Select

    Exit Sub

End Sub

Private Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SerieNFiscalOriginal_Validate

    'Verifica se foi preenchido
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then Exit Sub

    If SerieNFiscalOriginal.Text = SerieNFiscalOriginal.List(SerieNFiscalOriginal.ListIndex) Then Exit Sub

    If Len(Trim(SerieNFiscalOriginal.Text)) > 3 Then Error 30640

    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(SerieNFiscalOriginal)
    If lErro <> SUCESSO And lErro <> 12253 Then Error 55533

    If lErro = 12253 Then Error 55534

    Exit Sub

Erro_SerieNFiscalOriginal_Validate:

    Cancel = True


    Select Case Err

        Case 30640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_MAIOR", Err)

        Case 55533

        Case 55534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, SerieNFiscalOriginal.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166308)

    End Select

    Exit Sub

End Sub

Private Sub SerieNFOriginalLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    objSerie.sSerie = SerieNFiscalOriginal.Text

    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieNFOriginal)

End Sub

Private Sub Transportadora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Transportadora_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TransportadoraLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As Collection

    If Len(Trim(Transportadora.Text)) <> 0 Then objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)

    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

End Sub

Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not objNFiscal Is Nothing Then

        'Lê NFiscal no BD
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then Error 30636

        'Se NFiscal existe
        If lErro <> 31442 Then

            If objNFiscal.iTipoNFiscal <> DOCINFO_NRCC And objNFiscal.iTipoNFiscal <> DOCINFO_NRCP Then Error 30637

            'Preenche_Tela
            lErro = Preenche_Tela(objNFiscal)
            If lErro <> SUCESSO Then Error 30638

        Else
            'Se não existe
            Error 30639

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 30636, 30638

        Case 30637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_NAO_RECEBCLI", Err, objNFiscal.iTipoNFiscal)

        Case 30639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_NAO_CADASTRADO", Err, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166309)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEntrada_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntrada_DownClick

    'Verifica preenchimento de Data Entrada
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 30510

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case Err

        Case 30510

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166310)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntrada_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntrada_UpClick

    'Verifica preenchimento de Data Entrada
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 30511

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case Err

        Case 30511

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166311)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30656

    If objControl.Name = "Produto" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = False

        Else
            objControl.Enabled = True

        End If

    ElseIf objControl.Name = "UnidadeMed" Then

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
            objControl.Enabled = False

        Else
            objControl.Enabled = True

            objProduto.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 30658

            If lErro = 28030 Then gError 30659

            objClasseUM.iClasse = objProduto.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError 30660

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            'Limpar as Unidades utilizadas anteriormente
            UnidadeMed.Clear

            For Each objUnidadeDeMedida In colSiglas
                UnidadeMed.AddItem objUnidadeDeMedida.sSigla
            Next

            'Tento selecionar na Combo a Unidade anterior
            If UnidadeMed.ListCount <> 0 Then

                For iIndice = 0 To UnidadeMed.ListCount - 1

                    If UnidadeMed.List(iIndice) = sUnidadeMed Then
                        UnidadeMed.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If

        End If

'distribuicao
'    ElseIf objControl.Name = "Quantidade" Or objControl.Name = "Almoxarifado" Or objControl.Name = "ValorUnitario" Or objControl.Name = "DescricaoItem" Or objControl.Name = "Ccl" Or objControl.Name = "PercentDesc" Or objControl.Name = "Desconto" Then
    ElseIf objControl.Name = "ValorUnitario" Or objControl.Name = "DescricaoItem" Or objControl.Name = "Ccl" Or objControl.Name = "PercentDesc" Or objControl.Name = "Desconto" Then
        
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
            objControl.Enabled = False

        Else
            objControl.Enabled = True

        End If
    
    ElseIf objControl.Name = "Quantidade" Then
    
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If
    End If
        
    'distribuicao
    lErro = gobjDistribuicao.Rotina_Grid_Enable_Dist(iLinha, objControl, iLocalChamada)
    If lErro <> SUCESSO Then gError 89598
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 30656, 30658, 30659, 30660, 89598

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166312)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        If objGridInt.objGrid Is GridItens Then

            Select Case objGridInt.objGrid.Col

'distribuicao
'                'Célula Almoxarifado
'                Case iGrid_Almoxarifado_Col
'                    lErro = Saida_Celula_Almoxarifado(objGridInt)
'                    If lErro <> SUCESSO Then gError 30663

                'Célula Valor Unitário
                Case iGrid_ValorUnitario_Col
                    lErro = Saida_Celula_ValorUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 30664

                'Célula Produto
                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 30665

                'Célula Quantidade
                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 30666

                'Célula Unidade de Medida
                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 30667

                Case iGrid_PercDesc_Col
                    lErro = Saida_Celula_PercentDesc(objGridInt)
                    If lErro <> SUCESSO Then gError 28796

                Case iGrid_Desconto_Col
                    lErro = Saida_Celula_Desconto(objGridInt)
                    If lErro <> SUCESSO Then gError 28797
                
                'Célula Descricao
                Case iGrid_Descricao_Col
                    lErro = Saida_Celula_DescricaoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 49421

                'Célula Ccl
                Case iGrid_Ccl_Col
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 49382

            End Select

        'distribuicao
        ElseIf objGridInt.objGrid.Name = GridDist.Name Then

            lErro = gobjDistribuicao.Saida_Celula_Dist()
            If lErro <> SUCESSO Then gError 89599

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 30669

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 28796, 28797, 30663, 30664, 30665, 30666, 30667, 49382, 49421, 89599

        Case 30669
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166313)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Ccl que está deixando ser a corrente

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl
    
    'Verifica se Ccl foi preenchido
    If Len(Trim(Ccl.ClipText)) > 0 Then

        'Critica o Ccl
        lErro = CF("Ccl_Critica", Ccl, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then Error 49383

        If lErro = 5703 Then Error 49384

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49422
    
    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err

    Select Case Err

        Case 49383
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49384
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
        
        Case 49422
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166314)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49423

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = Err

    Select Case Err

        Case 49423
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 166315)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim iPossuiGrade As Integer

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica preenchimento do Produto
    If Len(Trim(Produto.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 30670

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            If lErro = 25041 Then Error 30671
            
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then Error 30672
    
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 30673

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = Err

    Select Case Err

        Case 30670, 30672, 30673
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30671
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166316)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long
'Preenche os dados do Produto da linha do grid selecionada

Dim lErro As Long
Dim iAlmoxarifadoPadrao As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iPossuiGrade As Integer
Dim sProdutoPai As String
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProduto As String
Dim objRomaneioGrade As New ClassRomaneioGrade
Dim iIndice As Integer
Dim objItemNF As ClassItemNF

On Error GoTo Erro_ProdutoLinha_Preenche

    If objProduto.iGerencial And Len(Trim(objProduto.sGrade)) = 0 Then gError 86296
    
    If Len(Trim(objProduto.sGrade)) > 0 Then iPossuiGrade = MARCADO

    If iPossuiGrade = DESMARCADO Then
    
        If Grid_Possui_Grade Then
        
            'Busca, caso exista, o produto pai de grade o prod em questão
            lErro = CF("Produto_Le_PaiGrade", objProduto, sProdutoPai)
            If lErro <> SUCESSO Then gError 86327
            
            'Se o produto tem um pai de grade
            If Len(Trim(sProdutoPai)) > 0 Then
                'Verifica se seu pai aparece no grid
                For iIndice = 1 To gobjNFiscal.ColItensNF.Count
                    'Se aparecer ==> erro
                    If gobjNFiscal.ColItensNF(iIndice).sProduto = sProdutoPai Then gError 86328
                
                Next
            
            End If
            
        End If
    Else
        'Verifica se há filhos válidos com a grade preenchida
        lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
        If lErro <> SUCESSO Then gError 86329
        
        'Se nao existir, erro
        If colItensRomaneioGrade.Count = 0 Then gError 86330
        
        'Para cada filho de grade do produto
        For Each objItensRomaneio In colItensRomaneioGrade
            'Verifica se ele já aparece no grid
            For iIndice = 1 To gobjNFiscal.ColItensNF.Count
                'Se aparecer ==> Erro
                If gobjNFiscal.ColItensNF(iIndice).sProduto = objItensRomaneio.sProduto Then gError 86331
            Next
        Next
    
    End If
            
    Set objItemNF = New ClassItemNF
    
    objItemNF.iPossuiGrade = iPossuiGrade

        
    objItemNF.sProduto = objProduto.sCodigo
    objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
    objItemNF.iItem = GridItens.Row
    objItemNF.lNumIntDoc = 0
    objItemNF.sDescricaoItem = objProduto.sDescricao
                
    If objItemNF.iPossuiGrade = MARCADO Then
        Set objRomaneioGrade = New ClassRomaneioGrade
        
        objRomaneioGrade.sNomeTela = Me.Name
        
        Set objRomaneioGrade.objObjetoTela = objItemNF
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
        If giRetornoTela <> vbOK Then gError 86310

        
    End If
    
    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda

    'Descricao
    GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
    
        gobjNFiscal.ColItensNF.Add1 objItemNF
        gobjNFiscal.ColItensNF(GridItens.Row).sUMEstoque = objProduto.sSiglaUMEstoque
        gobjNFiscal.ColItensNF(GridItens.Row).iItem = GridItens.Row
       
       If iPossuiGrade = MARCADO Then
        
            '************** GRADE ************
            gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade = MARCADO
                       
            Set gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade = objItemNF.colItensRomaneioGrade
            
            GridItens.TextMatrix(GridItens.Row, 0) = "# " & GridItens.TextMatrix(GridItens.Row, 0)
                   
            Call Atualiza_Grid_Itens(objItemNF)
            
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(objItemNF)
            
        End If
        
    End If

    Call Calcula_Valores
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 30317, 30318, 86310, 86327, 86329

        Case 30319
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case 86296
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 86328
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_GRID", gErr, Trim(sProdutoPai), Trim(Produto.Text))
                    
        Case 86330
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
        
        Case 86331
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, Trim(objProduto.sCodigo), Trim(gobjNFiscal.ColItensNF(iIndice).sProduto))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166317)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade que está deixando ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
'distribuicao
Dim dQuantidadeAnterior As Double
Dim dQuantidadeAtual As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'distribuicao
    dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    dQuantidadeAtual = StrParaDbl(Quantidade.Text)
    'fim  distribuicao

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 30457

        dQuantidade = CDbl(Quantidade.Text)

        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)
        
    End If

    'inicio distribuicao
    If dQuantidadeAnterior <> dQuantidadeAtual Then
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
        
        'altera as quantidades no tab de distribuicao
        lErro = gobjDistribuicao.Distribuicao_Processa()
        If lErro <> SUCESSO Then gError 89600
        
    End If
    'fim distribuicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30474

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores()
    If lErro <> SUCESSO Then gError 55512

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 30457, 30474, 55512, 89600
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166318)

    End Select

    Exit Function

End Function
  
Private Function Saida_Celula_ValorUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula valor unitário que está deixando ser a corrente

Dim lErro As Long
Dim dValorUnitario As Double

On Error GoTo Erro_Saida_Celula_ValorUnitario

    Set objGridInt.objControle = ValorUnitario

    'Se estiver preenchido
    If Len(Trim(ValorUnitario.ClipText)) > 0 Then
        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorUnitario.Text)
        If lErro <> SUCESSO Then Error 30681

        dValorUnitario = CDbl(ValorUnitario.Text)

        'Coloca o valor Formatado na tela
        ValorUnitario.Text = Format(dValorUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 30682

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores()
    If lErro <> SUCESSO Then Error 55523

    Saida_Celula_ValorUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorUnitario:

    Saida_Celula_ValorUnitario = Err

    Select Case Err

        Case 30681, 30682, 55523
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166319)

    End Select

    Exit Function

End Function

Private Function SubTotal_Calcula() As Long
'Soma a coluna de Valor Total e acumula em SubTotal

Dim lErro As Long
Dim dSubTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_SubTotal_Calcula

    For iIndice = 1 To objGrid.iLinhasExistentes

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col))) <> 0 Then
            dSubTotal = dSubTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col))

        End If

    Next

    SubTotal.Caption = Format(CStr(dSubTotal), "Standard")

    'Calcula o Total
    lErro = Total_Calcula()
    If lErro <> SUCESSO Then Error 30687

    SubTotal_Calcula = SUCESSO

    Exit Function

Erro_SubTotal_Calcula:

    SubTotal_Calcula = Err

    Select Case Err

        Case 30687

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166320)

    End Select

    Exit Function

End Function

Private Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)

    Call Valor_Saida(ValorFrete)

End Sub

Private Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)

    Call Valor_Saida(ValorSeguro)

End Sub

Private Sub ValorDespesas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)

    Call Valor_Saida(ValorDespesas)

End Sub

Private Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)

    Call Valor_Saida(ValorDesconto)

End Sub

Private Sub Valor_Saida(objControle As Object)

Dim lErro As Long

On Error GoTo Erro_Valor_Saida

    'Verifica se foi preenchido
    If Len(Trim(objControle.Text)) <> 0 Then

        'Critica se é Valor não negativo
        lErro = Valor_NaoNegativo_Critica(objControle.Text)
        If lErro <> SUCESSO Then Error 30688

        objControle.Text = Format(objControle.Text, "Fixed")

    End If

    'Calcula o Total
    lErro = Total_Calcula()
    If lErro <> SUCESSO Then Error 30689

    Exit Sub

Erro_Valor_Saida:

    Select Case Err

        Case 30688, 30689
            objControle.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166321)

    End Select

    Exit Sub

End Sub

Private Function Total_Calcula() As Long
'Calcula o Total

Dim dTotal As Double

    'Adiciona o SubTotal caso esteja preenchido
    If Len(Trim(SubTotal.Caption)) <> 0 And IsNumeric(SubTotal.Caption) Then dTotal = dTotal + CDbl(SubTotal.Caption)

    'Adiciona o Valor do Frete caso esteja preenchido
    If Len(Trim(ValorFrete.Text)) <> 0 And IsNumeric(ValorFrete.Text) Then dTotal = dTotal + CDbl(ValorFrete.Text)

    'Adiciona o Valor das Despesas caso esteja preenchido
    If Len(Trim(ValorDespesas.Text)) <> 0 And IsNumeric(ValorDespesas.Text) Then dTotal = dTotal + CDbl(ValorDespesas.Text)

    'Adiciona o Valor do Seguro caso esteja preenchido
    If Len(Trim(ValorSeguro.Text)) <> 0 And IsNumeric(ValorSeguro.Text) Then dTotal = dTotal + CDbl(ValorSeguro.Text)

    'Subtrai o Desconto caso esteja preenchido
    If Len(Trim(ValorDesconto.Text)) <> 0 And IsNumeric(ValorDesconto.Text) Then dTotal = dTotal - CDbl(ValorDesconto.Text)

    If Len(Trim(IPIValor1.Text)) > 0 And IsNumeric(IPIValor1.Text) Then dTotal = dTotal + CDbl(IPIValor1.Text)
    
    Total.Caption = Format(CStr(dTotal), "Standard")
    
    Total_Calcula = SUCESSO

End Function

Private Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 30690

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 30691

        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then Error 30692

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 30693

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True


    Select Case Err

        Case 30690, 30691

        Case 30692
            'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            Else
                'Segura o foco

            End If

        Case 30693
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166322)

    End Select

    Exit Sub

End Sub

Private Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'Verifica se foi preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO Then Error 30498

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case Err

        Case 30498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", Err, PlacaUF.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166323)

    End Select

    Exit Sub

End Sub

Private Sub PesoBruto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoBruto_Validate

    'Verifica se foi preenchido
    If Len(Trim(PesoBruto.Text)) = 0 Then Exit Sub

    lErro = Valor_NaoNegativo_Critica(PesoBruto.Text)
    If lErro <> SUCESSO Then Error 30694

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True


    Select Case Err

        Case 30694

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166324)

    End Select

    Exit Sub

End Sub

Private Sub PesoLiquido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoLiquido_Validate

    'Verifica se foi preenchido
    If Len(Trim(PesoLiquido.Text)) = 0 Then Exit Sub

    lErro = Valor_NaoNegativo_Critica(PesoLiquido.Text)
    If lErro <> SUCESSO Then Error 30695

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True


    Select Case Err

        Case 30695

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166325)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_RecebMaterialC()

Dim lErro As Long
Dim iIndice As Integer

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)
    
    Set gobjNFiscal = New ClassNFiscal
    
    For iIndice = 1 To objGrid.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next

    'Limpa o Grid
    Call Grid_Limpa(objGrid)
    
    'distribuicao
    Call gobjDistribuicao.Limpa_Tela_Distribuicao
    
    NumRecebimento.Caption = ""
    
    'Limpa o Label's
    SubTotal.Caption = ""
    Total.Caption = ""

    'Limpa e desseleciona a Combo Série
    Serie.Text = ""
    Serie.ListIndex = -1
    SerieNFiscalOriginal.Text = ""
    SerieNFiscalOriginal.ListIndex = -1

    'Desseleciona as combos Transportadora e PlacaUF
    Transportadora.ListIndex = -1
    Transportadora.Text = ""
    PlacaUF.ListIndex = -1
    PlacaUF.Text = ""

    'Incluído por Luiz Nogueira em 21/08/03
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)

    'Limpa as ComboBoxes
    Filial.Clear

    'Preenche a DataEntrada com a Data Atual
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True

    NFiscalPropria.Value = False
    NFiscalCliente.Value = False

    NFiscalOriginal.Text = ""
    Emitente.Value = True
 
    iAlterado = 0

End Sub

Private Sub ValorUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VolumeEspecie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeEspecie_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie, "AVISO_CRIAR_VOLUMEESPECIE")
    If lErro <> SUCESSO Then gError 102432
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102432
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166326)

    End Select

End Sub

Private Sub VolumeMarca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeMarca_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca, "AVISO_CRIAR_VOLUMEMARCA")
    If lErro <> SUCESSO Then gError 102433
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102433
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166327)

    End Select

End Sub

Private Sub VolumeNumero_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VolumeQuant_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Opcao_Click()

   'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        Select Case iFrameAtual
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_RECEBIMENTO_MATERIAL_CLIENTE_DADOS_PRINCIPAIS
                
            Case TAB_Itens
                Parent.HelpContextID = IDH_RECEBIMENTO_MATERIAL_CLIENTE_ITENS
            
            Case TAB_Complemento
                Parent.HelpContextID = IDH_RECEBIMENTO_MATERIAL_CLIENTE_COMPLEMENTO
                        
        End Select

    End If

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub ValorUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub ValorUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ValorUnitario
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentes As Integer
'distribuicao
Dim iItemAtual As Integer

On Error GoTo Erro_GridItens_KeyDown

    iLinhasExistentes = objGrid.iLinhasExistentes
    
    'distribuicao
    iItemAtual = GridItens.Row

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    If objGrid.iLinhasExistentes < iLinhasExistentes Then
    
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
        gobjNFiscal.ColItensNF.Remove GridItens.Row
        
        For iLinhasExistentes = 1 To objGrid.iLinhasExistentes
            If gobjNFiscal.ColItensNF(iLinhasExistentes).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iLinhasExistentes, 0) = "# " & iLinhasExistentes
            Else
                GridItens.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
            End If
            
        Next
        GridItens.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
        
        lErro = SubTotal_Calcula()
        If lErro <> SUCESSO Then gError 55524
    
        'distribuicao
        lErro = gobjDistribuicao.Exclusao_Item_GridDist(iItemAtual)
        If lErro <> SUCESSO Then gError 89603
    
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 55524, 89603

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166328)

    End Select
    
    Exit Sub

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Public Sub GridDist_Click()
'distribuicao
    
    Call gobjDistribuicao.GridDist_Click

End Sub

Public Sub GridDist_EnterCell()
'distribuicao
    
    Call gobjDistribuicao.GridDist_EnterCell

End Sub

Public Sub GridDist_GotFocus()
'distribuicao
    
    Call gobjDistribuicao.GridDist_GotFocus

End Sub

Public Sub GridDist_KeyPress(KeyAscii As Integer)
'distribuicao
    
    Call gobjDistribuicao.GridDist_KeyPress(KeyAscii)

End Sub

Public Sub GridDist_LeaveCell()
'distribuicao
    
    Call gobjDistribuicao.GridDist_LeaveCell

End Sub

Public Sub GridDist_Validate(Cancel As Boolean)
'distribuicao
    
    Call gobjDistribuicao.GridDist_Validate(Cancel)
    
End Sub

Public Sub GridDist_RowColChange()
'distribuicao
    
    Call gobjDistribuicao.GridDist_RowColChange

End Sub

Public Sub GridDist_KeyDown(KeyCode As Integer, Shift As Integer)
'distribuicao
    
    Call gobjDistribuicao.GridDist_KeyDown(KeyCode, Shift)
    
End Sub

Public Sub GridDist_Scroll()
'distribuicao
    
    Call gobjDistribuicao.GridDist_Scroll


End Sub

Function Gravar_Registro() As Long
'Valida os dados da tela para gravação

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As ClassProduto
Dim objNFiscal As New ClassNFiscal
Dim objNFiscalOriginal As New ClassNFiscal
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Chama Verifica_Preenchimento
    lErro = Verifica_Preenchimento()
    If lErro <> SUCESSO Then gError 30706

    'Verifica se algum Item no Grid
    If objGrid.iLinhasExistentes = 0 Then gError 30707

    'Para cada Linha do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 30708

        'Verifica se DescricaoItem foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Descricao_Col))) = 0 Then gError 49381

        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30910

        Set objProduto = New ClassProduto

        objProduto.sCodigo = sProdutoFormatado

        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 30911

        'Se não achou o Produto --> erro
        If lErro = 28030 Then gError 30912

        'Verifica se o Valor Unitário foi preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 30710

    Next

    'Verifica se Total é não negativo
    If Not (CDbl(Total.Caption) >= 0) Then gError 30711

    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 30714

    'distribuicao
    lErro = gobjDistribuicao.Move_GridDist_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 89601

    'Se SerieNFOriginal e NFiscalOriginal estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) <> 0 And Len(Trim(NFiscalOriginal.Text)) <> 0 Then

        objNFiscalOriginal.sSerie = SerieNFiscalOriginal.Text
        objNFiscalOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)

        'Lê NFiscal Interna Saida Nao Dev
        lErro = CF("NFiscalInternaSaidaNaoDev_Le_Numero", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 30765 Then gError 30712


        'Se não achou a NFiscal Interna Saida Nao Dev --> erro
        If lErro = 30765 Then gError 30713

        objNFiscal.lNumIntNotaOriginal = objNFiscalOriginal.lNumIntDoc

    End If

    'Chama RecebMaterialC_Grava
    lErro = CF("RecebMaterialC_Grava", objNFiscal)
    If lErro <> SUCESSO Then gError 30715

    GL_objMDIForm.MousePointer = vbDefault
    
    If Len(Trim(NumRecebimento.Caption)) = 0 Then vbMsg = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_RECEBIMENTO_GRAVADO", objNFiscal.lNumRecebimento)
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 30706, 30712, 30714, 30715, 30910, 30911, 89601

        Case 30707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSRECEB_NAO_INFORMADOS", gErr)

        Case 30708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADEITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 30709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADOITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 30710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIOITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 30711
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_RECEB_NEGATIVO", gErr)

        Case 30713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA", gErr, objNFiscalOriginal.sSerie, objNFiscalOriginal.lNumNotaFiscal)

        Case 30912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 49381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAOITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166329)

    End Select

    Exit Function

End Function

Private Function Verifica_Preenchimento() As Long
'Verifica preenchimento dos campos principais da tela

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer

On Error GoTo Erro_Verifica_Preenchimento

    'Se uma delas SerieNFiscalOriginal OU NFiscalOriginal está-->
    '--> mas a outra não ==> ERRO
    If (Len(Trim(SerieNFiscalOriginal.Text)) <> 0 And Len(Trim(NFiscalOriginal.Text)) = 0) Or (Len(Trim(SerieNFiscalOriginal.Text)) = 0 And Len(Trim(NFiscalOriginal.Text)) <> 0) Then Error 30698

    'Verifica se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then Error 30699

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Error 30700

    'Verifica se a DataEntrada foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then Error 30701

    'Verifica se foi selecionado algum Tipo de Nota fiscal
    If Not NFiscalCliente.Value And Not NFiscalPropria.Value Then Error 30702
    
    If NFiscalCliente.Value = True Then
    
        'Verifica se a Série foi preenchida
        If Len(Trim(Serie.Text)) = 0 Then Error 30703

        'Verifica se a Nota Fiscal foi preenchida
        If Len(Trim(NFiscal.Text)) = 0 Then Error 30704

        'Verifica se a Série de NotaFiscal Propria está Cadastrada no BD
        If NFiscalPropria.Value Then

            For iIndice = 0 To Serie.ListCount - 1
    
                If Serie.Text = Serie.List(iIndice) Then
                    iAchou = 1
                    Exit For
                End If
            Next

            If iAchou = 0 Then Error 30705

        End If
    
    End If
    
    'Verifica se o PesoBruto é maior que PesoLiq
    If Len(Trim(PesoLiquido.Text)) <> 0 And Len(Trim(PesoBruto.Text)) <> 0 Then

        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then Error 30771

    End If

    Verifica_Preenchimento = SUCESSO

    Exit Function

Erro_Verifica_Preenchimento:

    Verifica_Preenchimento = Err

    Select Case Err

        Case 30698
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NUMERO_ORIGINAL_FALTANDO", Err)

        Case 30699
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 30700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 30701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", Err)

        Case 30702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_INFORMADO", Err)

        Case 30703
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)

        Case 30704
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", Err)

        Case 30705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)

        Case 30771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESOBRUTO_MENOR_PESOLIQ", Err, CDbl(PesoBruto.Text), CDbl(PesoLiquido.Text))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166330)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long
Dim sUMAnterior As String

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

'inicio distribuicao
    'recolhe a UM anteriormente escolhida
    sUMAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    'coloca no grid a UM atual selecionda
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = UnidadeMed.Text
    
    If sUMAnterior <> UnidadeMed.Text Then
        
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade <> MARCADO Then
        
            'Tenta fazer a distribuição automatica p\ o item de acordo com a nova UM
            lErro = gobjDistribuicao.Distribuicao_Processa()
            If lErro <> SUCESSO Then gError 89602
        Else
            lErro = gobjDistribuicao.Distribuicao_Processa_Grade()
            If lErro <> SUCESSO Then gError 89602
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
        
        End If
                
    End If
    
'fim distribuicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30766

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 30766, 89602
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166331)

    End Select

    Exit Function

End Function

Private Function Coloca_Dados_Tela(objNFiscal As ClassNFiscal, objNFiscalOriginal As ClassNFiscal) As Long
'Mostra os dados passados em objNFiscal na tela

Dim lErro As Long
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Coloca_Dados_Tela

    NFiscalPropria.Value = False
    NFiscalCliente.Value = False

    'Coloca os Dados na Tela
    'Lê o NomeReduzido do Cliente no BD
    'Se o NumIntNotaOriginal for Positivo, preenche Tela com -->
    '--> Nota Fiscal Original
    If objNFiscal.lNumIntNotaOriginal > 0 Then
        NFiscalOriginal.Text = CStr(objNFiscalOriginal.lNumNotaFiscal)
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
    Else
        NFiscalOriginal.Text = ""
        SerieNFiscalOriginal.Text = ""
    End If
 
    NumRecebimento.Caption = objNFiscal.lNumRecebimento
    
    Cliente.Text = CStr(objNFiscal.lCliente)

    Call Cliente_Validate(bCancel)

    Filial.Text = CStr(objNFiscal.iFilialCli)

    Call Filial_Validate(bSGECancelDummy)

    If objNFiscal.dtDataEntrada <> DATA_NULA Then
        DataEntrada.PromptInclude = False
        DataEntrada.Text = Format(objNFiscal.dtDataEntrada, "dd/mm/yy")
        DataEntrada.PromptInclude = True

    Else
        DataEntrada.PromptInclude = False
        DataEntrada.Text = ""
        DataEntrada.PromptInclude = True

    End If

'horaentrada
    HoraEntrada.PromptInclude = False
    'este teste está correto
    If objNFiscal.dtDataEntrada <> DATA_NULA Then HoraEntrada.Text = Format(objNFiscal.dtHoraEntrada, "hh:mm:ss")
    HoraEntrada.PromptInclude = True

    If objNFiscal.iTipoNFiscal = DOCINFO_NRCP Then
    
        NFiscalPropria.Value = True

    Else
    
        NFiscalCliente.Value = True

    End If

    If Len(Trim(objNFiscal.sSerie)) > 0 Then
        Serie.Text = objNFiscal.sSerie
    Else
        Serie.Text = ""
    End If
    
    If objNFiscal.lNumNotaFiscal = 0 Then
        NFiscal.Text = ""
    Else
        NFiscal.Text = CStr(objNFiscal.lNumNotaFiscal)
    End If
    
    'Preenche o Tab Complemento

    'Seleciona a Transportadora através do Código no ItemData
    Transportadora.Text = ""
    For iIndice = 0 To Transportadora.ListCount - 1
        If Transportadora.ItemData(iIndice) = objNFiscal.iCodTransportadora Then
            Transportadora.ListIndex = iIndice
            Exit For

        End If
    Next

    Placa.Text = objNFiscal.sPlaca
    PlacaUF.Text = objNFiscal.sPlacaUF
    VolumeQuant.Text = CStr(objNFiscal.lVolumeQuant)
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a espécie dos volumes do pedido
    If objNFiscal.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objNFiscal.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    Else
        VolumeEspecie.Text = ""
    End If
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a marca dos volumes do pedido
    If objNFiscal.lVolumeMarca > 0 Then
        VolumeMarca.Text = objNFiscal.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    Else
        VolumeMarca.Text = ""
    End If
    
    Mensagem.Text = objNFiscal.sMensagemNota
    PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.iFreteRespons = FRETE_EMITENTE Then
        Emitente.Value = True
    Else
        Destinatario.Value = True
    End If

    VolumeNumero.Text = objNFiscal.sVolumeNumero

    lErro = Preenche_GridItens(objNFiscal.ColItensNF)
    If lErro <> SUCESSO Then gError 30621

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 89605

    SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")
    ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    IPIValor1.Text = Format(objNFiscal.dValorTotal - objNFiscal.dValorProdutos + objNFiscal.dValorDesconto - objNFiscal.dValorSeguro - objNFiscal.dValorFrete - objNFiscal.dValorOutrasDespesas, "Standard")

    lErro = Total_Calcula()
    If lErro <> SUCESSO Then gError 30622

    Coloca_Dados_Tela = SUCESSO

    Exit Function

Erro_Coloca_Dados_Tela:

    Coloca_Dados_Tela = gErr

    Select Case gErr

        Case 30621, 30622, 89605

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166332)

    End Select

End Function
Private Sub PercentDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub PercentDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub PercentDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = PercentDesc
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc.Text)) > 0 Then
    
        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then Error 30644

        dPercentDesc = CDbl(PercentDesc.Text)

        lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
        
        If lTamanho > 0 Then dPercentDescAnterior = StrParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

        If dPercentDesc <> dPercentDescAnterior Then

            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then Error 30645

            PercentDesc.Text = Format(dPercentDesc, "Fixed")

        End If
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
     If lErro <> SUCESSO Then Error 30776

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores()
    If lErro <> SUCESSO Then Error 55517

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = Err

    Select Case Err

        Case 30644, 30776, 55517
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166333)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dPrecoTotal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim iDescontoAlterado As Integer

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto

    iDescontoAlterado = False

'mario
    'Veifica o preenchimento de Desconto
    If Len(Trim(Desconto.ClipText)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then Error 30641

        dDesconto = CDbl(Desconto.Text)
            
        If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) <> dDesconto Then iDescontoAlterado = True
        
        If iDescontoAlterado = True Then
        
            dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
            dPrecoTotal = dQuantidade * dPrecoUnitario
        
            If dPrecoTotal > 0 Then
        
                If dDesconto >= dPrecoTotal Then Error 30642
        
                dPercentDesc = dDesconto / dPrecoTotal
        
                GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")
        
            End If
            
        End If

    Else
'mario
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) <> 0 Then

            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = ""
        
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 30643

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores()
    If lErro <> SUCESSO Then Error 55518

    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = Err

    Select Case Err

        Case 30641, 30643, 55518
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", Err, GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166334)

    End Select

    Exit Function

End Function

Private Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = dValorUnitario * dQuantidade

    'Se o Percentual Desconto estiver preenchido
    If dPercentDesc > 0 Then

        'Testa se o desconto está preenchido
        If dDesconto = 0 Then
            dPercDesc2 = 0
        Else
            'Calcula o Percentual em cima dos valores passados
            dPercDesc2 = dDesconto / dValorTotal
            dPercDesc2 = CDbl(Format(dPercDesc2, "0.0000"))
        End If
        'se os percentuais (passado e calculado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = dPercentDesc * dValorTotal

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

Private Function Calcula_Valores() As Long
'recalcula os valores de desconto, percentual de desconto e valor total

Dim sProduto As String
Dim lErro As Long
Dim lTamanho As Long
Dim dPercentDesc As Double
Dim dValorUnitario As Double
Dim dDesconto As Double
Dim dValorReal As Double
Dim dQuantidade As Double

On Error GoTo Erro_Calcula_Valores

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))

    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
    If dQuantidade = 0 Or Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) = 0 Then

        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorTotal_Col) = ""
        
    Else

        lTamanho = Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)))

        If lTamanho > 0 Then
            dPercentDesc = PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
        Else
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        End If

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dValorUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))) > 0 Then dDesconto = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dValorUnitario, dPercentDesc, dDesconto, dValorReal)

        'Coloca o Desconto calculado na tela
        If dDesconto > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
        Else
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        End If

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorTotal_Col) = Format(dValorReal, "Standard")

    End If

    lErro = SubTotal_Calcula()
    If lErro <> SUCESSO Then Error 55525

    Calcula_Valores = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valores:

    Calcula_Valores = Err
    
    Select Case Err

        Case 55525

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166335)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RECEBIMENTO_MATERIAL_CLIENTE_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Recebimento de Material do Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RecebMaterialC"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is SerieNFiscalOriginal Then
            Call SerieNFOriginalLabel_Click
        ElseIf Me.ActiveControl Is NFiscalOriginal Then
            Call NFOriginalLabel_Click
        ElseIf Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NumRecebimento Then
            Call LabelRecebimento_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
'distribuicao
        ElseIf Me.ActiveControl Is gobjDistribuicao.AlmoxDist Then
            Call gobjDistribuicao.BotaoLocalizacaoDist_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        
        End If
    
    End If

End Sub

Private Sub VolumeQuant_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)

End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub SerieLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieLabel, Source, X, Y)
End Sub

Private Sub SerieLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieLabel, Button, Shift, X, Y)
End Sub

Private Sub NFiscalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalLabel, Source, X, Y)
End Sub

Private Sub NFiscalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalLabel, Button, Shift, X, Y)
End Sub

Private Sub NumRecebimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumRecebimento, Source, X, Y)
End Sub

Private Sub NumRecebimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumRecebimento, Button, Shift, X, Y)
End Sub

Private Sub LabelRecebimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRecebimento, Source, X, Y)
End Sub

Private Sub LabelRecebimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRecebimento, Button, Shift, X, Y)
End Sub

Private Sub SerieNFOriginalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieNFOriginalLabel, Source, X, Y)
End Sub

Private Sub SerieNFOriginalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieNFOriginalLabel, Button, Shift, X, Y)
End Sub

Private Sub NFOriginalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFOriginalLabel, Source, X, Y)
End Sub

Private Sub NFOriginalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFOriginalLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub SubTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SubTotal, Source, X, Y)
End Sub

Private Sub SubTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SubTotal, Button, Shift, X, Y)
End Sub

Private Sub Total_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Total, Source, X, Y)
End Sub

Private Sub Total_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Total, Button, Shift, X, Y)
End Sub

Private Sub LabelIPIValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelIPIValor, Source, X, Y)
End Sub

Private Sub LabelIPIValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelIPIValor, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Function Grid_Possui_Grade() As Boolean

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim iIndice As Integer

    For iIndice = 1 To gobjNFiscal.ColItensNF.Count
        If gobjNFiscal.ColItensNF(iIndice).iPossuiGrade = MARCADO Then
            Grid_Possui_Grade = True
            Exit Function
        End If
    Next
    
    Grid_Possui_Grade = False
        
    Exit Function
    
End Function

Public Sub BotaoGrade_Click()

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemNF As ClassItemNF

On Error GoTo Erro_BotaoGrade_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        Set objItemNF = gobjNFiscal.ColItensNF(GridItens.Row)
        
        If objItemNF.iPossuiGrade = MARCADO Then
            
            Set objRomaneioGrade = New ClassRomaneioGrade
            
            objRomaneioGrade.sNomeTela = Me.Name
            Set objRomaneioGrade.objObjetoTela = objItemNF
                        
            Call gobjDistribuicao.Move_DistribuicaoGrade_Memoria(objItemNF)
            
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
        
            Call Atualiza_Grid_Itens(objItemNF)
            
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(objItemNF)
        
            Call Calcula_Valores
        
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166336)
            
    End Select
    
    Exit Sub

End Sub


Sub Atualiza_Grid_Itens(objItemNF As ClassItemNF)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
    
    For Each objItemRomaneioGrade In objItemNF.colItensRomaneioGrade
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
    Next

    GridItens.TextMatrix(objItemNF.iItem, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    objItemNF.dQuantidade = dQuantidade
    
    Exit Sub

End Sub

Function Move_ItensGrade_Tela(colItensRomaneio As Collection, colItensRomaneioTela As Collection, Optional bTrazPedido As Boolean = False) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objReservaItemTela As ClassReservaItem
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long

On Error GoTo Erro_Move_ItensGrade_Tela

    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradeTela In colItensRomaneioTela
                    
        Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
        objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto
        objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
        objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
        objItemRomaneioGrade.dQuantPV = objItemRomaneioGradeTela.dQuantidade - objItemRomaneioGradeTela.dQuantCancelada - objItemRomaneioGradeTela.dQuantFaturada
        objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
        objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
        objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade - objItemRomaneioGradeTela.dQuantCancelada - objItemRomaneioGradeTela.dQuantFaturada
        objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
        objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
        If bTrazPedido Then
            objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntDoc
        Else
            objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
            objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntItemPV
        End If
                            
                            
        colItensRomaneio.Add objItemRomaneioGrade
    
        'Transfere as informações de Localização
        Set objItemRomaneioGrade.colLocalizacao = New Collection
            
        For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
            
            If objReservaItemTela.iAlmoxarifado > 0 Then
            
                objAlmoxarifado.iCodigo = objReservaItemTela.iAlmoxarifado
                            
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> 25056 And lErro <> SUCESSO Then gError 94331
                If lErro = 25056 Then gError 94332
                
                objAlmoxarifado.sNomeReduzido = objAlmoxarifado.sNomeReduzido
            
            Else
                objAlmoxarifado.sNomeReduzido = objReservaItemTela.sAlmoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 33145
        
                'Se não encontrou o Nome Reduzido do Almoxarifado
                If lErro <> SUCESSO Then gError 33146
            
                objReservaItemTela.iAlmoxarifado = objAlmoxarifado.iCodigo
            End If
            
            
            If objAlmoxarifado.iFilialEmpresa = giFilialEmpresa Then
            
                Set objReservaItem = New ClassReservaItem
                
                objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
                objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
                objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
                objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
                objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
                objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
                objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                
                objItemRomaneioGrade.colLocalizacao.Add objReservaItem
            End If
        Next
    
    Next
    
    Move_ItensGrade_Tela = SUCESSO
    
    Exit Function
    
Erro_Move_ItensGrade_Tela:

    Move_ItensGrade_Tela = gErr

    Select Case gErr
            
        Case 33145, 94331

        Case 33146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 94332
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objReservaItemTela.iAlmoxarifado)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166337)

    End Select
    
    Exit Function

End Function

Function Transfere_Dados_ItensRomaneio(colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio existente do BD (Produtos Filhos do Produto passado)
    For Each objItemRomaneioGrade In colItensRomaneio
        'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
        For Each objItemRomaneioGradeTela In colItensRomaneioTela
            'Se encontrou o Item
            If objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto Then
                'Transfere as informações vindas da tela chamadora para essa tela
                objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
                objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
                objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
                objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
                objItemRomaneioGrade.dQuantFaturada = objItemRomaneioGradeTela.dQuantFaturada
                objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
                objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
                objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
                objItemRomaneioGrade.dQuantCancelada = objItemRomaneioGradeTela.dQuantCancelada
                objItemRomaneioGrade.dQuantPV = objItemRomaneioGradeTela.dQuantPV
                objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntItemPV
                objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
                
                'Transfere as informações de Localização
                Set objItemRomaneioGrade.colLocalizacao = New Collection
                    
                For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
                    
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
                    objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
                    objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
                    objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
                    objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
                    objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
                    objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                    
                    objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                    
                Next
                            
            End If
        
        Next
    Next

    Exit Function

End Function

Public Sub ProdutoAlmoxDist_Change()
'distribuicao

    Call gobjDistribuicao.ProdutoAlmoxDist_Change

End Sub

Public Sub ProdutoAlmoxDist_GotFocus()
'distribuicao

    Call gobjDistribuicao.ProdutoAlmoxDist_GotFocus

End Sub

Public Sub ProdutoAlmoxDist_KeyPress(KeyAscii As Integer)
'distribuicao

    Call gobjDistribuicao.ProdutoAlmoxDist_KeyPress(KeyAscii)

End Sub

Public Sub ProdutoAlmoxDist_Validate(Cancel As Boolean)
'distribuicao

    Call gobjDistribuicao.ProdutoAlmoxDist_Validate(Cancel)

End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134032

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166338)

    End Select
    
    Exit Sub

End Sub


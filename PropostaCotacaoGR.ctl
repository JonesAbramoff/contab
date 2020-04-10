VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PropostaCotacao 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9435
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4845
      Index           =   2
      Left            =   90
      TabIndex        =   25
      Top             =   825
      Visible         =   0   'False
      Width           =   9270
      Begin VB.TextBox TextServico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   216
         MaxLength       =   100
         TabIndex        =   29
         Top             =   510
         Width           =   5196
      End
      Begin VB.TextBox TextOrigem 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5700
         MaxLength       =   50
         TabIndex        =   28
         Top             =   408
         Width           =   1530
      End
      Begin VB.TextBox TextDestino 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7308
         MaxLength       =   50
         TabIndex        =   27
         Top             =   408
         Width           =   1530
      End
      Begin VB.TextBox TextObsServico 
         Height          =   765
         Left            =   1515
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   3600
         Width           =   7650
      End
      Begin MSFlexGridLib.MSFlexGrid GridDestOrigem 
         Height          =   3345
         Left            =   75
         TabIndex        =   30
         Top             =   210
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5900
         _Version        =   393216
         Rows            =   13
      End
      Begin MSComCtl2.UpDown UpDownDataInicio 
         Height          =   300
         Left            =   2640
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4470
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskDataInicio 
         Height          =   300
         Left            =   1530
         TabIndex        =   32
         Top             =   4470
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Previsão Início:"
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
         Index           =   1
         Left            =   105
         TabIndex        =   34
         Top             =   4455
         Width           =   1365
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   4
         Left            =   375
         TabIndex        =   33
         Top             =   3615
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4830
      Index           =   4
      Left            =   36
      TabIndex        =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   9270
      Begin VB.ComboBox ComboCondPagto 
         Height          =   315
         Left            =   2550
         TabIndex        =   23
         Top             =   60
         Width           =   2400
      End
      Begin VB.Frame Frame2 
         Caption         =   "Serviços"
         Height          =   2844
         Index           =   4
         Left            =   270
         TabIndex        =   10
         Top             =   360
         Width           =   8808
         Begin VB.TextBox TextDescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1692
            MaxLength       =   50
            TabIndex        =   21
            Top             =   930
            Width           =   2892
         End
         Begin MSMask.MaskEdBox UFDestino 
            Height          =   225
            Left            =   4305
            TabIndex        =   11
            Top             =   1695
            Width           =   525
            _ExtentX        =   926
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
         Begin MSMask.MaskEdBox Destino 
            Height          =   225
            Left            =   2775
            TabIndex        =   12
            Top             =   2085
            Width           =   1050
            _ExtentX        =   1852
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
         Begin MSMask.MaskEdBox Origem 
            Height          =   225
            Left            =   1635
            TabIndex        =   13
            Top             =   1605
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSMask.MaskEdBox UFOrigem 
            Height          =   225
            Left            =   3030
            TabIndex        =   14
            Top             =   1695
            Width           =   480
            _ExtentX        =   847
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
         Begin MSMask.MaskEdBox MaskQuantidade 
            Height          =   225
            Left            =   4650
            TabIndex        =   15
            Top             =   930
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSMask.MaskEdBox MaskPreco 
            Height          =   225
            Left            =   7125
            TabIndex        =   16
            Top             =   930
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSMask.MaskEdBox MaskPedagio 
            Height          =   225
            Left            =   6960
            TabIndex        =   17
            Top             =   1455
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSMask.MaskEdBox MaskAdValoren 
            Height          =   225
            Left            =   5790
            TabIndex        =   18
            Top             =   1485
            Width           =   1155
            _ExtentX        =   2037
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskProduto 
            Height          =   225
            Left            =   270
            TabIndex        =   19
            Top             =   870
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskPrecoUnit 
            Height          =   228
            Left            =   5868
            TabIndex        =   20
            Top             =   936
            Width           =   1152
            _ExtentX        =   2037
            _ExtentY        =   423
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
         Begin MSFlexGridLib.MSFlexGrid GridServicos 
            Height          =   2385
            Left            =   150
            TabIndex        =   22
            Top             =   285
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   4207
            _Version        =   393216
            Rows            =   8
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Resultado"
         Height          =   1095
         Index           =   1
         Left            =   225
         TabIndex        =   3
         Top             =   3255
         Width           =   8820
         Begin VB.TextBox TextObsResultado 
            Height          =   300
            Left            =   1860
            MaxLength       =   255
            TabIndex        =   6
            Top             =   660
            Width           =   6105
         End
         Begin VB.ComboBox ComboSituacaoResultado 
            Height          =   315
            ItemData        =   "PropostaCotacaoGR.ctx":0000
            Left            =   1860
            List            =   "PropostaCotacaoGR.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   225
            Width           =   1740
         End
         Begin VB.ComboBox ComboJustificativa 
            Height          =   315
            ItemData        =   "PropostaCotacaoGR.ctx":002C
            Left            =   6120
            List            =   "PropostaCotacaoGR.ctx":002E
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   1848
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Situação:"
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
            Height          =   192
            Index           =   11
            Left            =   960
            TabIndex        =   9
            Top             =   288
            Width           =   828
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Justificativa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   9
            Left            =   4944
            TabIndex        =   8
            Top             =   300
            Width           =   1092
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
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
            Height          =   192
            Index           =   10
            Left            =   684
            TabIndex        =   7
            Top             =   672
            Width           =   1092
         End
      End
      Begin VB.CommandButton BotaoOrigemDestino 
         Caption         =   "Origem/Destino"
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
         Left            =   7095
         TabIndex        =   2
         Top             =   4455
         Width           =   1965
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
         Height          =   345
         Left            =   5025
         TabIndex        =   1
         Top             =   4455
         Width           =   1965
      End
      Begin VB.Label LabelCondPagto 
         AutoSize        =   -1  'True
         Caption         =   "Condição de Pagamento:"
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
         TabIndex        =   24
         Top             =   90
         Width           =   2145
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4845
      HelpContextID   =   1
      Index           =   3
      Left            =   36
      TabIndex        =   35
      Top             =   765
      Visible         =   0   'False
      Width           =   9270
      Begin VB.CheckBox CheckDescarga 
         Caption         =   "Descarga:"
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
         Left            =   5115
         TabIndex        =   50
         Top             =   585
         Width           =   1200
      End
      Begin VB.CheckBox CheckCarga 
         Caption         =   "Carga:"
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
         Left            =   1275
         TabIndex        =   49
         Top             =   585
         Width           =   885
      End
      Begin VB.ComboBox ComboDescarga 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PropostaCotacaoGR.ctx":0030
         Left            =   6375
         List            =   "PropostaCotacaoGR.ctx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   600
         Width           =   1605
      End
      Begin VB.ComboBox ComboDesova 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PropostaCotacaoGR.ctx":0062
         Left            =   6375
         List            =   "PropostaCotacaoGR.ctx":006C
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1050
         Width           =   1605
      End
      Begin VB.ComboBox ComboTipoEmbalagem 
         Height          =   315
         Left            =   2265
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   195
         Width           =   1605
      End
      Begin VB.CheckBox CheckCargaSolta 
         Caption         =   "Carga Solta"
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
         Left            =   5115
         TabIndex        =   45
         Top             =   1485
         Width           =   1380
      End
      Begin VB.Frame FrameTipoConteiner 
         Caption         =   "Container"
         Height          =   2475
         Left            =   1245
         TabIndex        =   41
         Top             =   2325
         Width           =   6795
         Begin VB.ComboBox ComboTipoContainer 
            Height          =   315
            ItemData        =   "PropostaCotacaoGR.ctx":0094
            Left            =   645
            List            =   "PropostaCotacaoGR.ctx":0096
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   735
            Width           =   4590
         End
         Begin MSMask.MaskEdBox MaskQuantCtr 
            Height          =   225
            Left            =   5400
            TabIndex        =   42
            Top             =   630
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridContainer 
            Height          =   2145
            Left            =   120
            TabIndex        =   44
            Top             =   255
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   3784
            _Version        =   393216
         End
      End
      Begin VB.TextBox TextDescCSolta 
         Height          =   285
         Left            =   2265
         MaxLength       =   255
         TabIndex        =   40
         Top             =   1950
         Width           =   5790
      End
      Begin VB.ComboBox ComboOva 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PropostaCotacaoGR.ctx":0098
         Left            =   2265
         List            =   "PropostaCotacaoGR.ctx":00A2
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1050
         Width           =   1620
      End
      Begin VB.ComboBox ComboCarga 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PropostaCotacaoGR.ctx":00CA
         Left            =   2265
         List            =   "PropostaCotacaoGR.ctx":00D4
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   600
         Width           =   1620
      End
      Begin VB.CheckBox CheckOva 
         Caption         =   "Ova:"
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
         Left            =   1275
         TabIndex        =   37
         Top             =   1035
         Width           =   735
      End
      Begin VB.CheckBox CheckDesova 
         Caption         =   "Desova:"
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
         Left            =   5115
         TabIndex        =   36
         Top             =   1035
         Width           =   1035
      End
      Begin MSMask.MaskEdBox MaskQuantAjudantes 
         Height          =   315
         Left            =   6360
         TabIndex        =   51
         Top             =   165
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskValorMerc 
         Height          =   315
         Left            =   2265
         TabIndex        =   52
         Top             =   1485
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Embalagem:"
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
         Left            =   420
         TabIndex        =   56
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label LabelDescCSolta 
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
         Left            =   1230
         TabIndex        =   55
         Top             =   1980
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ajudantes:"
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
         Left            =   5370
         TabIndex        =   54
         Top             =   225
         Width           =   915
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Mercadoria:"
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
         Index           =   5
         Left            =   645
         TabIndex        =   53
         Top             =   1545
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4845
      Index           =   1
      Left            =   90
      TabIndex        =   57
      Top             =   825
      Width           =   9270
      Begin VB.TextBox TextObservacao 
         Height          =   285
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   79
         Top             =   4470
         Visible         =   0   'False
         Width           =   7740
      End
      Begin VB.Frame Frame2 
         Caption         =   "Vendas"
         Height          =   660
         Index           =   0
         Left            =   60
         TabIndex        =   74
         Top             =   3675
         Width           =   9135
         Begin VB.TextBox TextIndicacao 
            Height          =   315
            Left            =   5550
            MaxLength       =   50
            TabIndex        =   75
            Top             =   240
            Width           =   3435
         End
         Begin MSMask.MaskEdBox MaskVendedor 
            Height          =   315
            Left            =   1740
            TabIndex        =   76
            Top             =   240
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Indicação:"
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
            Index           =   7
            Left            =   4590
            TabIndex        =   78
            Top             =   285
            Width           =   915
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
            Left            =   795
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   77
            Top             =   285
            Width           =   885
         End
      End
      Begin VB.ComboBox ComboTipoOperacao 
         Height          =   315
         ItemData        =   "PropostaCotacaoGR.ctx":00FC
         Left            =   5280
         List            =   "PropostaCotacaoGR.ctx":0109
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   120
         Width           =   1560
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cliente"
         Height          =   3090
         Left            =   60
         TabIndex        =   58
         Top             =   525
         Width           =   9135
         Begin VB.Frame Frame2 
            Caption         =   "Envio da Cotação"
            Height          =   660
            Index           =   2
            Left            =   105
            TabIndex        =   66
            Top             =   2310
            Width           =   8940
            Begin VB.ComboBox ComboEnvioCotacao 
               Height          =   315
               ItemData        =   "PropostaCotacaoGR.ctx":0136
               Left            =   840
               List            =   "PropostaCotacaoGR.ctx":0146
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   270
               Width           =   1260
            End
            Begin VB.TextBox TextEnvioComplemento 
               Height          =   315
               Left            =   3495
               MaxLength       =   100
               TabIndex        =   67
               Top             =   240
               Width           =   5280
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Meio:"
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
               Index           =   2
               Left            =   300
               TabIndex        =   70
               Top             =   300
               Width           =   480
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Complemento:"
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
               Left            =   2235
               TabIndex        =   69
               Top             =   300
               Width           =   1200
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Contatos"
            Height          =   1560
            Left            =   90
            TabIndex        =   59
            Top             =   645
            Width           =   8940
            Begin VB.TextBox TextSetor 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   2580
               MaxLength       =   50
               TabIndex        =   64
               Top             =   570
               Width           =   1725
            End
            Begin VB.TextBox TextContato 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   525
               MaxLength       =   50
               TabIndex        =   63
               Top             =   540
               Width           =   2025
            End
            Begin VB.TextBox TextFax 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   5550
               MaxLength       =   18
               TabIndex        =   62
               Top             =   615
               Width           =   1170
            End
            Begin VB.TextBox TextTelefone 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   4320
               MaxLength       =   18
               TabIndex        =   61
               Top             =   630
               Width           =   1170
            End
            Begin VB.TextBox TextEmail 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   6720
               MaxLength       =   50
               TabIndex        =   60
               Top             =   630
               Width           =   1815
            End
            Begin MSFlexGridLib.MSFlexGrid GridContatos 
               Height          =   1215
               Left            =   105
               TabIndex        =   65
               Top             =   240
               Width           =   8730
               _ExtentX        =   15399
               _ExtentY        =   2143
               _Version        =   393216
            End
         End
         Begin MSMask.MaskEdBox MaskCliente 
            Height          =   300
            Left            =   1125
            TabIndex        =   71
            Top             =   300
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label LabelCliente 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
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
            Left            =   270
            TabIndex        =   72
            Top             =   315
            Width           =   795
         End
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   2385
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   150
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskData 
         Height          =   300
         Left            =   1275
         TabIndex        =   81
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   84
         Top             =   4485
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label label1 
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
         Index           =   0
         Left            =   735
         TabIndex        =   83
         Top             =   195
         Width           =   480
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Operação:"
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
         Index           =   8
         Left            =   3600
         TabIndex        =   82
         Top             =   150
         Width           =   1605
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7320
      ScaleHeight     =   495
      ScaleWidth      =   1980
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   15
      Width           =   2040
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1515
         Picture         =   "PropostaCotacaoGR.ctx":0168
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1005
         Picture         =   "PropostaCotacaoGR.ctx":02E6
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   525
         Picture         =   "PropostaCotacaoGR.ctx":0818
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "PropostaCotacaoGR.ctx":09A2
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   270
      Left            =   2196
      Picture         =   "PropostaCotacaoGR.ctx":0AFC
      Style           =   1  'Graphical
      TabIndex        =   85
      ToolTipText     =   "Numeração Automática"
      Top             =   75
      Width           =   315
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5235
      Left            =   30
      TabIndex        =   91
      Top             =   450
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   9234
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Serviços"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Carga"
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
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   285
      Left            =   1395
      TabIndex        =   92
      Top             =   75
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label NumeroLabel 
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
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   615
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   93
      Top             =   75
      Width           =   735
   End
End
Attribute VB_Name = "PropostaCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Tulio
'Inicio: 21/01
'Termino: 30/01

'Definicoes do Grid de Destino Origem
Dim objGridDestOrigem As New AdmGrid

Dim iGrid_Servico_Col As Integer
Dim iGrid_Origem_Col As Integer
Dim iGrid_Destino_Col As Integer

'Definicoes do Grid de Containers
Dim objGridContainer As New AdmGrid

Dim iGrid_Container_Col As Integer
Dim iGrid_Quantidade_Col As Integer

'Definicoes do Grid de Contatos
Dim objGridContatos As New AdmGrid

Dim iGrid_Contato_Col As Integer
Dim iGrid_Setor_Col As Integer
Dim iGrid_Telefone_Col As Integer
Dim iGrid_Fax_Col As Integer
Dim iGrid_Email_Col As Integer

'Definicoes do Grid de Serviços
Dim objGridServicos As New AdmGrid

Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoItem_Col As Integer
Dim iGrid_QuantServico_Col As Integer
Dim iGrid_PrecoUnitario_Col As Integer
Dim iGrid_Preco_Col As Integer
Dim iGrid_AdValoren_Col As Integer
Dim iGrid_Pedagio_Col As Integer
Dim iGrid_OrigemServ_Col As Integer
Dim iGrid_UFOrigemServ_Col As Integer
Dim iGrid_DestinoServ_Col As Integer
Dim iGrid_UFDestinoServ_Col As Integer

Public iFrameAtual As Integer
Public iAlterado As Integer
Public iVendedorAlterado As Integer

'Definições dos TABS
Private Const TAB_Identificacao = 1
Private Const TAB_Servico = 2
Private Const TAB_Carga = 3
'Private Const TAB_CondComerciais = 4

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos browser
Private WithEvents objEventoCotacao As AdmEvento
Attribute objEventoCotacao.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoOrigem As AdmEvento
Attribute objEventoOrigem.VB_VarHelpID = -1
Private WithEvents objEventoDestino As AdmEvento
Attribute objEventoDestino.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long
Dim iIndiceFrame As Integer
Dim colEmbalagens As New Collection
Dim colContainers As New Collection
Dim objTipoEmb As ClassTipoEmbalagem
Dim objTipoContainer As ClassTipoContainer

On Error GoTo Erro_Form_Load

    iFrameAtual = TAB_Identificacao
    Set objEventoOrigem = New AdmEvento
    Set objEventoDestino = New AdmEvento
    Set objEventoCotacao = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoServico = New AdmEvento
    
    'Torna os frames invisiveis a fim de só tornar
    'visivel o frame correspondente e deixa o
    'primeiro visivel
    Frame1(TAB_Identificacao).Visible = True
    For iIndiceFrame = TAB_Servico To TAB_Carga
        Frame1(iIndiceFrame).Visible = False
    Next
    
    'Mascara o produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MaskProduto)
    If lErro <> SUCESSO Then gError 95342
    
    'Executa inicializacao do GridContatos
    lErro = Inicializa_Grid_Contatos(objGridContatos)
    If lErro <> SUCESSO Then gError 97343

'    'Executa inicializacao do GridServicos
'    lErro = Inicializa_Grid_Servicos(objGridServicos)
'    If lErro <> 0 Then gError 97344
    
    'Executa inicializacao do GridDestOrigem
    lErro = Inicializa_Grid_DestOrigem(objGridDestOrigem)
    If lErro <> 0 Then gError 97345

    'Executa inicializacao do GridContainer
    lErro = Inicializa_Grid_Container(objGridContainer)
    If lErro <> 0 Then gError 97346

    'Inicializa os Eventos
    Set objEventoCotacao = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoServico = New AdmEvento
    
    'Inicializa data com a data atual
    MaskData.PromptInclude = False
    MaskData.Text = Format(gdtDataHoje, "dd/mm/yy")
    MaskData.PromptInclude = True
    
    lErro = Carrega_CondicaoPagamento
    If lErro <> SUCESSO Then gError 97122
    
    'Inicializando combo TipoEmbalagem com
    'as embalagens lidas da tabela
    lErro = CF("TipoEmbalagem_Le_Todos", colEmbalagens)
    If lErro <> SUCESSO Then gError 97123

    For Each objTipoEmb In colEmbalagens
        'Adiciona o item na combo de tipo embalagem e preenche o itemdata
        ComboTipoEmbalagem.AddItem objTipoEmb.sDescricao
        ComboTipoEmbalagem.ItemData(ComboTipoEmbalagem.NewIndex) = objTipoEmb.iTipo

    Next
    
    'Inicializando combo TipoContainer com
    'os containers lidos da tabela
    lErro = CF("TipoContainer_Le_Todos", colContainers)
    If lErro <> SUCESSO Then gError 97124
    
    For Each objTipoContainer In colContainers
        'Adiciona o item na combo de tipo container e preenche o itemdata
        ComboTipoContainer.AddItem objTipoContainer.iTipo & SEPARADOR & objTipoContainer.sDescricao
        ComboTipoContainer.ItemData(ComboTipoContainer.NewIndex) = objTipoContainer.iTipo

    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 97122, 97123, 97124, 97343, 97344, 97345, 97346
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Traz_Cotacao_Tela(objCotacao As ClassCotacaoGR) As Long
'Traz os dados da Cotacao para a tela

Dim lErro As Long

On Error GoTo Erro_Traz_Cotacao_Tela

    'Le a Cotacao
    lErro = CF("Cotacao_LeGR", objCotacao)
    If lErro <> SUCESSO And lErro <> 97170 Then gError 97171

    'Nao achou a cotacao
    If lErro <> SUCESSO Then gError 97189

    'Le a tabela de Contato referentes a cotacao lida anteriormente
    lErro = CF("Contato_Le_Cotacao", objCotacao, objCotacao.colContato)
    If lErro <> SUCESSO And lErro <> 97194 Then gError 97196

    'Se nao achou, erro de integridade, pois
    'a cotacao deve possuir ao menos 1 contato associado
    If lErro <> SUCESSO Then gError 97381

    'Le a tabela de CotacaoDestinoOrigem referentes a cotacao lida anteriormente
    lErro = CF("CotacaoOrigemDestino_Le_Cotacao", objCotacao, objCotacao.colCotacaoOrigemDestino)
    If lErro <> SUCESSO And lErro <> 97183 Then gError 97198

    'Se nao achou, erro de integridade, pois
    'a cotacao deve possuir ao menos 1 Origem/Destino
    If lErro <> SUCESSO Then gError 97382

    'Le a tabela de CotacaoServico referentes a cotacao lida anteriormente
    lErro = CF("CotacaoServico_Le_Cotacao", objCotacao, objCotacao.colCotacaoServico)
    If lErro <> SUCESSO And lErro <> 97188 Then gError 97200

    'Nao precisa tratar pois uma cotacao pode ser
    'gravada sem produto (servico) preenchido
    'If lErro <> SUCESSO Then gError 97383

    'Se a carga for solta
    If objCotacao.iCargaSolta <> CARGA_SOLTA Then
    
        'Le a tabela de CotacaoContainer referentes a cotacao lida anteriormente
        lErro = CF("CotacaoContainer_Le_Cotacao", objCotacao, objCotacao.colCotacaoContainer)
        If lErro <> SUCESSO And lErro <> 97178 Then gError 97202
    
        'Se nao achou, erro de integridade, pois
        'a cotacao deve possuir ao menos 1 container associado
        'quando a carga nao eh solta
        If lErro <> SUCESSO Then gError 97384
    
    End If

    'Move os dados pra tela
    lErro = Move_Dados_Tela(objCotacao)
    If lErro <> SUCESSO Then gError 97347
    
    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Traz_Cotacao_Tela = SUCESSO

    Exit Function

Erro_Traz_Cotacao_Tela:

    Traz_Cotacao_Tela = gErr

    Select Case gErr

        Case 97171, 97196, 97198, 97200, 97202, 97347
        
        Case 97189
            Call Rotina_Erro(vbOKOnly, "ERRO_COTACAO_NAO_CADASTRADA2", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
            
        Case 97381
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_COTACAO_CONTATO", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
        
        Case 97382
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_COTACAO_COTACAODESTINOORIGEM", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
            
'        Case 97383
'            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_COTACAO_COTACAOSERVICO", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
        
        Case 97384
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_COTACAO_COTACAOCONTAINER", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Move_Dados_Tela(objCotacao As ClassCotacaoGR) As Long
'Move os dados carregados em objCotacao para a tela

Dim lErro As Long

On Error GoTo Erro_Move_Dados_Tela
    
    Call Limpa_Tela_Cotacao
    
    MaskCodigo.Text = objCotacao.lCodigo
    
    'Traz as informacoes do primeiro TAB pra tela
    lErro = Move_Dados_Tela_PrimeiroTAB(objCotacao)
    If lErro <> SUCESSO Then gError 97360
    
    'Traz as informacoes do segundo TAB pra tela
    lErro = Move_Dados_Tela_SegundoTAB(objCotacao)
    If lErro <> SUCESSO Then gError 97361
    
    'Traz as informacoes do terceiro TAB pra tela
    lErro = Move_Dados_Tela_TerceiroTAB(objCotacao)
    If lErro <> SUCESSO Then gError 97362
    
    'Traz as informacoes do quarto TAB pra tela
    lErro = Move_Dados_Tela_QuartoTAB(objCotacao)
    If lErro <> SUCESSO Then gError 97363
        
    Move_Dados_Tela = SUCESSO
    
    Exit Function

Erro_Move_Dados_Tela:

    Move_Dados_Tela = gErr

    Select Case gErr

        Case 97360, 97361, 97362, 97363
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
   
End Function

Function Move_Dados_Tela_PrimeiroTAB(objCotacao As ClassCotacaoGR) As Long
'Move os dados carregados em objCotacao para o primeiro TAB da tela

Dim lErro As Long
    
On Error GoTo Erro_Move_Dados_Tela_PrimeiroTAB
    
    'Move a data da cotacao
    MaskData.PromptInclude = False
    MaskData.Text = Format(objCotacao.dtData, "dd/mm/yy")
    MaskData.PromptInclude = True
    
    'Move o TipoOperacao
    ComboTipoOperacao.ListIndex = objCotacao.iTipoOperacao
        
    'Move a empresa
    MaskCliente.Text = objCotacao.sCliente
    
    'Move o Meio de Envio
    ComboEnvioCotacao.ListIndex = objCotacao.iEnvio
    
    'Move o complemento do envio
    TextEnvioComplemento.Text = objCotacao.sEnvioComplemento
    
    'Move a observacao
    TextObservacao.Text = objCotacao.sObservacao
    
    'Se o vendedor foi gravado
    If objCotacao.iCodVendedor > 0 Then
    
        'Move o codigo do vendedor para a MaskEditBox "Vendedor" na tela
        MaskVendedor.Text = objCotacao.iCodVendedor
        
        'Chama o validate afim de recuperar o nome reduzido do vendedor
        'obs. O nome reduzido deve ser exibido na tela...
        Call MaskVendedor_Validate(bSGECancelDummy)
    
    Else
    'Se o vendedor nao foi gravado
        
        'A maskEdit "Vendedor" eh limpa
        MaskVendedor.Text = ""
    
    End If
    
    'Move a Indicacao
    TextIndicacao = objCotacao.sIndicacao
    
    'Carrega o grid de contatos
    lErro = Carrega_GridContatos(objCotacao)
    If lErro <> SUCESSO Then gError 97348

    Move_Dados_Tela_PrimeiroTAB = SUCESSO

    Exit Function

Erro_Move_Dados_Tela_PrimeiroTAB:

    Move_Dados_Tela_PrimeiroTAB = gErr
    
    Select Case gErr
    
        Case 97348
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function
    
End Function

Function Move_Dados_Tela_SegundoTAB(objCotacao As ClassCotacaoGR) As Long
'Move os dados carregados em objCotacao para o segundo TAB da tela

Dim lErro As Long
    
On Error GoTo Erro_Move_Dados_Tela_SegundoTAB
    'cyntia
    'Move a observacao Servico
    TextObsServico.Text = objCotacao.sObsDestOrigem
    
    'Se a data de previsao do inicio dos servicos foi gravada
    If objCotacao.dtDataPrevInicio <> DATA_NULA Then
        
        'Move a data de previsao do inicio dos servicos
        MaskDataInicio.PromptInclude = False
        MaskDataInicio.Text = Format(objCotacao.dtDataPrevInicio, "dd/mm/yy")
        MaskDataInicio.PromptInclude = True
        
    End If
    
    'Carrega o Grid de Origem/Destino
    lErro = Carrega_GridDestOrigem(objCotacao)
    If lErro <> SUCESSO Then gError 97349
    
    Move_Dados_Tela_SegundoTAB = SUCESSO

    Exit Function

Erro_Move_Dados_Tela_SegundoTAB:

    Move_Dados_Tela_SegundoTAB = gErr
    
    Select Case gErr
    
        Case 97349
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function
    
End Function

Function Move_Dados_Tela_TerceiroTAB(objCotacao As ClassCotacaoGR) As Long
'Move os dados carregados em objCotacao para o terceiro TAB da tela

Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_Move_Dados_Tela_TerceiroTAB

    'Varre a combo de tipoEmbalagem para ver qual delas
    'foi a gravada e atribuir o indice da mesma ao
    'listindex
    For iIndex = 0 To ComboTipoEmbalagem.ListCount - 1
        If ComboTipoEmbalagem.ItemData(iIndex) = objCotacao.iTipoEmbalagem Then
            ComboTipoEmbalagem.ListIndex = iIndex
            Exit For
        End If
    Next

    'Move o numero de ajudantes
    If objCotacao.iAjudantes > 0 Then MaskQuantAjudantes = objCotacao.iAjudantes

    'Se a cotacao foi gravada com a check "Carga" marcada
    If objCotacao.iCarga = INCLUI_CARGA Then

        'Marca a check "Carga" na tela
        CheckCarga.Value = vbChecked

        'Habilita a combo carga na tela
        ComboCarga.Enabled = True

        'Muda o listindex da combo carga, que coincide,
        'propositalmente com o campo gravado no BD
        ComboCarga.ListIndex = objCotacao.iCargaPorConta

    Else
    'Se a cotacao foi gravada com a check "Carga" desmarcada

        'Desmarca a check "Carga" na tela
        CheckCarga.Value = vbUnchecked

        'Desabilita a combo carga na tela
        ComboCarga.Enabled = False

        'Limpa a Combo
        ComboCarga.ListIndex = -1

    End If

    'Se a cotacao foi gravada com a check "DesCarga" marcada
    If objCotacao.iDesCarga = INCLUI_DESCARGA Then

        'Marca a check "DesCarga" na tela
        CheckDescarga.Value = vbChecked

        'Habilita a combo descarga na tela
        ComboDescarga.Enabled = True

        'Muda o listindex da combo descarga, que coincide,
        'propositalmente com o campo gravado no BD
        ComboDescarga.ListIndex = objCotacao.iDesCargaPorConta

    Else
    'Se a cotacao foi gravada com a check "DesCarga" desmarcada

        'Desmarca a check "DesCarga" na tela
        CheckDescarga.Value = vbUnchecked

        'Desabilita a combo descarga na tela
        ComboDescarga.Enabled = False

        'limpa a combo
        ComboDescarga.ListIndex = -1

    End If

    'Se a cotacao foi gravada com a check "Ova" marcada
    If objCotacao.iOva = INCLUI_OVA Then

        'Marca a check "Ova" na tela
        CheckOva.Value = vbChecked

        'Habilita a combo ova na tela
        ComboOva.Enabled = True

        'Muda o listindex da combo ova, que coincide,
        'propositalmente com o campo gravado no BD
        ComboOva.ListIndex = objCotacao.iOvaPorConta

    Else
    'Se a cotacao foi gravada com a check "Ova" desmarcada

        'Desmarca a check "Ova" na tela
        CheckOva.Value = vbUnchecked

        'Desabilita a combo ova na tela
        ComboOva.Enabled = False

        'limpa a combo
        ComboOva.ListIndex = -1

    End If


    'Se a cotacao foi gravada com a check "DesOva" marcada
    If objCotacao.iDesova = INCLUI_DESOVA Then

        'Marca a check "DesOva" na tela
        CheckDesova.Value = vbChecked

        'Habilita a combo DesOva na tela
        ComboDesova.Enabled = True

        'Muda o listindex da combo Desova, que coincide,
        'propositalmente com o campo gravado no BD
        ComboDesova.ListIndex = objCotacao.iDesovaPorConta

    Else
    'Se a cotacao foi gravada com a check "DesOva" desmarcada

        'Desmarca a check "DesOva" na tela
        CheckDesova.Value = vbUnchecked

        'Desabilita a combo desova na tela
        ComboDesova.Enabled = False

        'Limpa a Combo
        ComboDesova.ListIndex = -1

    End If

    'Se a cotacao foi gravada com a check "Carga Solta" marcada
    If objCotacao.iCargaSolta = CARGA_SOLTA Then

        'Marca a check "Carga Solta" na tela
        CheckCargaSolta.Value = vbChecked

    Else
    'Se a cotacao foi gravada com a check "Carga Solta" desmarcada

        'Desmarca a check "Carga Solta" na tela
        CheckCargaSolta.Value = vbUnchecked

    End If
    
    'Move a Descricao da Carga Solta
    TextDescCSolta.Text = objCotacao.sDescCargaSolta
    
    MaskValorMerc.Text = Format(objCotacao.dValorMerc, "standard")

    'Carrega o Grid de Origem/Destino
    lErro = Carrega_GridContainer(objCotacao)
    If lErro <> SUCESSO Then gError 97350

    Move_Dados_Tela_TerceiroTAB = SUCESSO

    Exit Function

Erro_Move_Dados_Tela_TerceiroTAB:

    Move_Dados_Tela_TerceiroTAB = gErr

    Select Case gErr

        Case 97350

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Move_Dados_Tela_QuartoTAB(objCotacao As ClassCotacaoGR) As Long
'Move os dados carregados em objCotacao para o quarto TAB da tela

Dim lErro As Long

On Error GoTo Erro_Move_Dados_Tela_QuartoTAB

    'Se a condicao de pagamento foi preenchida
    If objCotacao.iCondicaoPagto > 0 Then

        'Coloca o codigo da condicao na tela
        ComboCondPagto.Text = objCotacao.iCondicaoPagto

        'Chama o validate afim de recuperar o nome reduzido do vendedor
        'obs. O nome reduzido deve ser exibido na tela...
        Call ComboCondPagto_Validate(bSGECancelDummy)

    Else
    'Se a condicao de pagamento nao foi preenchida

        'Limpa a condicao de pagamento
        ComboCondPagto.ListIndex = -1

    End If

    'Move a situacao resultado
    'Obs. O valor gravado corresponde, propositalmente,
    'ao listindex da combo
    ComboSituacaoResultado.ListIndex = objCotacao.iSituacao

    'codigos serao definidos no futuro..
    'ComboJustificativa.ListIndex = objCotacao.iJustificativa

    'Move a Observacao Resultado
    TextObsResultado = objCotacao.sObsResultado

'    'Carrega o Grid de Origem/Destino
'    lErro = Carrega_GridServicos(objCotacao)
'    If lErro <> SUCESSO Then gError 97351

    Move_Dados_Tela_QuartoTAB = SUCESSO

    Exit Function

Erro_Move_Dados_Tela_QuartoTAB:

    Move_Dados_Tela_QuartoTAB = gErr

    Select Case gErr

        Case 97351

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Carrega_GridContatos(objCotacao As ClassCotacaoGR) As Long
'Preenche o Grid de contatos com o Conteudo do BD

Dim iLinha As Integer
Dim objContato As ClassContato

On Error GoTo Erro_Carrega_GridContatos

    'Limpa o Grid de Contato
    Call Grid_Limpa(objGridContatos)

    iLinha = 0

    'Preenche o grid com os objetos da coleção de contato
    For Each objContato In objCotacao.colContato

       iLinha = iLinha + 1

       GridContatos.TextMatrix(iLinha, iGrid_Contato_Col) = objContato.sContato
       GridContatos.TextMatrix(iLinha, iGrid_Setor_Col) = objContato.sSetor
       GridContatos.TextMatrix(iLinha, iGrid_Email_Col) = objContato.sEmail
       GridContatos.TextMatrix(iLinha, iGrid_Telefone_Col) = objContato.sTelefone
       GridContatos.TextMatrix(iLinha, iGrid_Fax_Col) = objContato.sFax

    Next

    objGridContatos.iLinhasExistentes = iLinha

    Exit Function
    
Erro_Carrega_GridContatos:

    Carrega_GridContatos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function
'
'Private Function Carrega_GridServicos(objCotacao As ClassCotacaoGR) As Long
''Preenche o Grid de Servicos com o Conteudo do BD
'
'Dim iLinha As Integer
'Dim objCotacaoServico As ClassCotacaoServico
'Dim objProduto As New ClassProduto
'Dim lErro As Long
'
''Flags q indicam se a quantidade e o preco unitario sao zero
'Dim bQtdEhZero As Boolean
'Dim bPrecUnitEhZero As Boolean
'Dim objOrigemDestino As New ClassOrigemDestino
'
'On Error GoTo Erro_Carrega_GridServicos
'
'    'Limpa o Grid de Servicos
'    Call Grid_Limpa(objGridServicos)
'
'    iLinha = 0
'
'    'Preenche o grid com os objetos da coleção de cotacaoservicos
'    For Each objCotacaoServico In objCotacao.colCotacaoServico
'
'        'Inicializando as booleans
'        bQtdEhZero = True
'        bPrecUnitEhZero = True
'
'        iLinha = iLinha + 1
'
'        MaskProduto.PromptInclude = False
'        MaskProduto.Text = objCotacaoServico.sProduto
'        MaskProduto.PromptInclude = True
'
'        'pega a descricao
'        lErro = Traz_Produto_Tela(objProduto)
'        If lErro <> SUCESSO Then gError 97386
'
'        GridServicos.TextMatrix(iLinha, iGrid_Produto_Col) = MaskProduto.Text
'        GridServicos.TextMatrix(iLinha, iGrid_DescricaoItem_Col) = objProduto.sDescricao
'
'        If objCotacaoServico.dQuantidade > 0 Then
'             GridServicos.TextMatrix(iLinha, iGrid_QuantServico_Col) = objCotacaoServico.dQuantidade
'             bQtdEhZero = False
'        End If
'
'        If objCotacaoServico.dPrecoUnitario > 0 Then
'             GridServicos.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(objCotacaoServico.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
'             bPrecUnitEhZero = False
'        End If
'
'        If bQtdEhZero = False And bPrecUnitEhZero = False Then
'             GridServicos.TextMatrix(iLinha, iGrid_Preco_Col) = Format(objCotacaoServico.dPrecoUnitario * objCotacaoServico.dQuantidade, gobjFAT.sFormatoPrecoUnitario)
'        End If
'
'        If objCotacaoServico.dAdValoren > 0 Then
'             GridServicos.TextMatrix(iLinha, iGrid_AdValoren_Col) = Format(objCotacaoServico.dAdValoren, "Percent")
'        End If
'
'        If objCotacaoServico.dPedagio > 0 Then
'             GridServicos.TextMatrix(iLinha, iGrid_Pedagio_Col) = Format(objCotacaoServico.dPedagio, "STANDARD")
'        End If
'
'        If objCotacaoServico.iDestino > 0 Then
'
'            objOrigemDestino.iCodigo = objCotacaoServico.iDestino
'
'            lErro = CF("OrigemDestino_Le", objOrigemDestino)
'            If lErro <> SUCESSO And lErro <> 96567 Then gError 99242
'
'            If lErro = 96567 Then gError 99243
'
'            GridServicos.TextMatrix(iLinha, iGrid_DestinoServ_Col) = objOrigemDestino.sOrigemDestino
'            GridServicos.TextMatrix(iLinha, iGrid_UFDestinoServ_Col) = objOrigemDestino.sUF
'
'        End If
'
'        If objCotacaoServico.iOrigem > 0 Then
'            objOrigemDestino.iCodigo = objCotacaoServico.iOrigem
'
'            lErro = CF("OrigemDestino_Le", objOrigemDestino)
'            If lErro <> SUCESSO And lErro <> 96567 Then gError 99244
'
'            If lErro = 96567 Then gError 99245
'
'            GridServicos.TextMatrix(iLinha, iGrid_OrigemServ_Col) = objOrigemDestino.sOrigemDestino
'            GridServicos.TextMatrix(iLinha, iGrid_UFOrigemServ_Col) = objOrigemDestino.sUF
'        End If
'
'    Next
'
'    objGridServicos.iLinhasExistentes = iLinha
'
'    Carrega_GridServicos = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_GridServicos:
'
'    Carrega_GridServicos = gErr
'
'    Select Case gErr
'
'        Case 97386, 99242, 99244
'
'        Case 99243
'            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_EXISTENTE", gErr, objOrigemDestino.iCodigo)
'
'        Case 99245
'            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_EXISTENTE", gErr, objOrigemDestino.iCodigo)
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_GridContainer(objCotacao As ClassCotacaoGR) As Long
'Preenche o Grid de Container com o Conteudo do BD

Dim iLinha As Integer
Dim objCotacaoContainer As ClassCotacaoContainer
Dim iIndex As Integer

On Error GoTo Erro_Carrega_GridContainer:

    'Limpa o Grid de Container
    Call Grid_Limpa(objGridContainer)

    iLinha = 0

    'Preenche o grid com os objetos da coleção de cotacaocontainer
    For Each objCotacaoContainer In objCotacao.colCotacaoContainer

       iLinha = iLinha + 1

       GridContainer.TextMatrix(iLinha, iGrid_Quantidade_Col) = objCotacaoContainer.iQuantidade

       'busca no itemdata pelo codigo do container
        For iIndex = 0 To ComboTipoContainer.ListCount - 1
            If Codigo_Extrai(ComboTipoContainer.List(iIndex)) = objCotacaoContainer.iTipoContainer Then
                GridContainer.TextMatrix(iLinha, iGrid_Container_Col) = ComboTipoContainer.List(iIndex)
                Exit For
            End If
        Next


    Next

    objGridContainer.iLinhasExistentes = iLinha

    Carrega_GridContainer = SUCESSO

    Exit Function

Erro_Carrega_GridContainer:

    Carrega_GridContainer = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Carrega_GridDestOrigem(objCotacao As ClassCotacaoGR)
'Preenche o Grid de Destino Origem com o Conteudo do BD

Dim iLinha As Integer
Dim objDestOrigem As ClassCotacaoOrigemDestino

On Error GoTo Erro_Carrega_GridDestOrigem

    'Limpa o Grid de Origem Destino
    Call Grid_Limpa(objGridDestOrigem)

    iLinha = 0

    'Preenche o grid com os objetos da coleção de DestOrigem
    For Each objDestOrigem In objCotacao.colCotacaoOrigemDestino

       iLinha = iLinha + 1

       GridDestOrigem.TextMatrix(iLinha, iGrid_Servico_Col) = objDestOrigem.sServico
       GridDestOrigem.TextMatrix(iLinha, iGrid_Origem_Col) = objDestOrigem.sOrigem
       GridDestOrigem.TextMatrix(iLinha, iGrid_Destino_Col) = objDestOrigem.sDestino
       
    Next

    objGridDestOrigem.iLinhasExistentes = iLinha

    Carrega_GridDestOrigem = SUCESSO
    
    Exit Function
    
Erro_Carrega_GridDestOrigem:
    
    Carrega_GridDestOrigem = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_DestOrigem(objGridInt As AdmGrid) As Long
'Inicializa o grid de Destino/Origem da tela

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_DestOrigem

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Origem")
    objGridInt.colColuna.Add ("Destino")

    'campos de edição do grid
    objGridInt.colCampo.Add (TextServico.Name)
    objGridInt.colCampo.Add (TextOrigem.Name)
    objGridInt.colCampo.Add (TextDestino.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Servico_Col = 1
    iGrid_Origem_Col = 2
    iGrid_Destino_Col = 3

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridDestOrigem

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_DESTORIGEM + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridDestOrigem.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_DestOrigem = SUCESSO

    Exit Function

Erro_Inicializa_Grid_DestOrigem:

    Inicializa_Grid_DestOrigem = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Container(objGridInt As AdmGrid) As Long
'Inicializa o grid de Container da tela

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Container

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Container")
    objGridInt.colColuna.Add ("Quantidade")

    'campos de edição do grid
    objGridInt.colCampo.Add (ComboTipoContainer.Name)
    objGridInt.colCampo.Add (MaskQuantCtr.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Container_Col = 1
    iGrid_Quantidade_Col = 2

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridContainer

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_CONTAINER + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridContainer.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Container = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Container:

    Inicializa_Grid_Container = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Contatos(objGridInt As AdmGrid) As Long
'Inicializa o grid de Contatos da tela

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Contatos

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Contato")
    objGridInt.colColuna.Add ("Setor")
    objGridInt.colColuna.Add ("Telefone")
    objGridInt.colColuna.Add ("Fax")
    objGridInt.colColuna.Add ("E-Mail")

    'campos de edição do grid
    objGridInt.colCampo.Add (TextContato.Name)
    objGridInt.colCampo.Add (TextSetor.Name)
    objGridInt.colCampo.Add (TextTelefone.Name)
    objGridInt.colCampo.Add (TextFax.Name)
    objGridInt.colCampo.Add (TextEmail.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Contato_Col = 1
    iGrid_Setor_Col = 2
    iGrid_Telefone_Col = 3
    iGrid_Fax_Col = 4
    iGrid_Email_Col = 5

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridContatos

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_CONTATOS + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridContatos.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contatos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Contatos:

    Inicializa_Grid_Contatos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

'Private Function Inicializa_Grid_Servicos(objGridInt As AdmGrid) As Long
''Inicializa o grid de Servicos da tela
'
'Dim lErro As Long
'
'On Error GoTo Erro_Inicializa_Grid_Servicos
'
'    'Tela em questão
'    Set objGridInt.objForm = Me
'
'    'Titulos do grid
'    objGridInt.colColuna.Add ("")
'    objGridInt.colColuna.Add ("Serviço")
'    objGridInt.colColuna.Add ("Descrição")
'    objGridInt.colColuna.Add ("Quantidade")
'    objGridInt.colColuna.Add ("Preço Unitário")
'    objGridInt.colColuna.Add ("Preço")
'    objGridInt.colColuna.Add ("AdValoren")
'    objGridInt.colColuna.Add ("Pedágio")
'    objGridInt.colColuna.Add ("Origem")
'    objGridInt.colColuna.Add ("UF")
'    objGridInt.colColuna.Add ("Destino")
'    objGridInt.colColuna.Add ("UF")
'
'    'campos de edição do grid
'    objGridInt.colCampo.Add (MaskProduto.Name)
'    objGridInt.colCampo.Add (TextDescricaoItem.Name)
'    objGridInt.colCampo.Add (MaskQuantidade.Name)
'    objGridInt.colCampo.Add (MaskPrecoUnit.Name)
'    objGridInt.colCampo.Add (MaskPreco.Name)
'    objGridInt.colCampo.Add (MaskAdValoren.Name)
'    objGridInt.colCampo.Add (MaskPedagio.Name)
'    objGridInt.colCampo.Add (Origem.Name)
'    objGridInt.colCampo.Add (UFOrigem.Name)
'    objGridInt.colCampo.Add (Destino.Name)
'    objGridInt.colCampo.Add (UFDestino.Name)
'
'    'indica onde estao situadas as colunas do grid
'    iGrid_Produto_Col = 1
'    iGrid_DescricaoItem_Col = 2
'    iGrid_QuantServico_Col = 3
'    iGrid_PrecoUnitario_Col = 4
'    iGrid_Preco_Col = 5
'    iGrid_AdValoren_Col = 6
'    iGrid_Pedagio_Col = 7
'    iGrid_OrigemServ_Col = 8
'    iGrid_UFOrigemServ_Col = 9
'    iGrid_DestinoServ_Col = 10
'    iGrid_UFDestinoServ_Col = 11
'
'
'    'Relaciona com o grid correspondente na tela
'    objGridInt.objGrid = GridServicos
'
'    'Linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAX_SERVICOS + 1
'
'    'linhas visiveis do grid
'    objGridInt.iLinhasVisiveis = 8
'
'    'Largura da primeira coluna
'    objGridInt.objGrid.ColWidth(0) = 400
'
'    'Usado para que se possa utilizar a Rotina_Enable
'    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
'
'    'largura total do grid
'    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'
'    'Chama rotina de Inicialização do Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_Servicos = SUCESSO
'
'    Exit Function
'
'Erro_Inicializa_Grid_Servicos:
'
'    Inicializa_Grid_Servicos = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
Public Function TP_OrigemDestino_Le(objOrigemDestino As ClassOrigemDestino, sOrigemDestino As String, sUF As String) As Long
'Lê a Origem ou o Destino com Código ou NomeRed

Dim eTipoOrigemDestino As enumTipo
Dim lErro As Long

On Error GoTo TP_OrigemDestino_Le

    eTipoOrigemDestino = Tipo_OrigemDestino(sOrigemDestino)

    Select Case eTipoOrigemDestino

    Case TIPO_STRING
        
        objOrigemDestino.sOrigemDestino = sOrigemDestino
                    
        If Len(Trim(sUF)) <> 0 Then
        
            objOrigemDestino.sUF = sUF
            
            lErro = CF("OrigemDestino_Le_NomeUF", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96860 Then gError 99253
            
            'Não existe OrigemDestino com este Nome e UF
            If lErro = 96860 Then gError 99254
                        
        Else
                    
            lErro = CF("OrigemDestino_Le_Nome", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96864 Then gError 99230
            
            If lErro = 96864 Then gError 99231
            
        End If
        
    Case TIPO_CODIGO

        objOrigemDestino.iCodigo = StrParaInt(sOrigemDestino)
               
        lErro = CF("OrigemDestino_Le", objOrigemDestino)
        If lErro <> SUCESSO And lErro <> 96567 Then gError 99232
            
        If lErro = 96567 Then gError 99233
            
      
    Case TIPO_DECIMAL

        gError 99234

    Case TIPO_NAO_POSITIVO

        gError 99235

    End Select

    TP_OrigemDestino_Le = SUCESSO

    Exit Function

TP_OrigemDestino_Le:

    TP_OrigemDestino_Le = gErr

    Select Case gErr
        
        Case 99230, 99232, 99253, 99255 'Tratados nas rotinas chamadas
                
        Case 99231
            'Envia aviso que OrigemDestino não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ORIGEMDESTINO", objOrigemDestino.sOrigemDestino)
    
                If lErro = vbYes Then
                    'Chama tela de OrigemDestino
                    lErro = Chama_Tela("OrigemDestino", objOrigemDestino)
                End If
                
        Case 99233
            'Envia aviso que OrigemDestino não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ORIGEMDESTINO", objOrigemDestino.iCodigo)
    
                If lErro = vbYes Then
                    'Chama tela de OrigemDestino
                    lErro = Chama_Tela("OrigemDestino", objOrigemDestino)
                End If
        
        Case 99234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sOrigemDestino)

        Case 99235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sOrigemDestino)
        
        Case 99254
            'Envia aviso que OrigemDestino não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ORIGEMDESTINO_UF", objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)
    
                If lErro = vbYes Then
                    'Chama tela de OrigemDestino
                    lErro = Chama_Tela("OrigemDestino", objOrigemDestino)
                End If
        
        Case 99256
            'Envia aviso que OrigemDestino não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ORIGEMDESTINO_UF", objOrigemDestino.iCodigo, objOrigemDestino.sUF)
    
                If lErro = vbYes Then
                    'Chama tela de OrigemDestino
                    lErro = Chama_Tela("OrigemDestino", objOrigemDestino)
                End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function

End Function

Private Function Tipo_OrigemDestino(ByVal sText As String) As enumTipo

If Not IsNumeric(sText) Then
    Tipo_OrigemDestino = TIPO_STRING
ElseIf Int(CDbl(sText)) <> CDbl(sText) Then
    Tipo_OrigemDestino = TIPO_DECIMAL
ElseIf CDbl(sText) <= 0 Then
    Tipo_OrigemDestino = TIPO_NAO_POSITIVO
Else
    Tipo_OrigemDestino = TIPO_CODIGO
End If

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objCotacao As New ClassCotacaoGR
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    objCotacao.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objCotacao.lNumIntDoc <> 0 Then

        'Carrega o obj com dados de colcampovalor (so o codigo e a filial)
        objCotacao.lCodigo = colCampoValor.Item("Codigo").vValor
        objCotacao.iFilialEmpresa = giFilialEmpresa
       
        'Traz para tela os dados carregados...
        'O caso de nao encontrar nao foi tratado,
        'pois nesse caso a rotina "Traz_Cotacao_Tela"
        'é interrompida e nao traz os dados pra tela
        lErro = Traz_Cotacao_Tela(objCotacao)
        If lErro <> SUCESSO Then gError 97387
                                     
        iAlterado = 0

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 97387

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCotacao As New ClassCotacaoGR
Dim sObsDestOrigem1 As String
Dim sObsDestOrigem2 As String
Dim sObsDestOrigem3 As String
Dim sObsDestOrigem4 As String

On Error GoTo Erro_Tela_Extrai

    sTabela = "CotacaoGR"

    lErro = Move_Tela_Memoria(objCotacao)
    If lErro <> SUCESSO Then gError 97250
    
    'Cyntia
    sObsDestOrigem1 = Mid(objCotacao.sObsDestOrigem, 1, 250)
    sObsDestOrigem2 = Mid(objCotacao.sObsDestOrigem, 251, 250)
    sObsDestOrigem3 = Mid(objCotacao.sObsDestOrigem, 501, 250)
    sObsDestOrigem4 = Mid(objCotacao.sObsDestOrigem, 751, 250)

    'Preenche a coleção colCampoValor
    colCampoValor.Add "NumIntDoc", objCotacao.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "FilialEmpresa", objCotacao.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Codigo", objCotacao.lCodigo, 0, "Codigo"
    colCampoValor.Add "Data", objCotacao.dtData, 0, "Data"
    colCampoValor.Add "TipoOperacao", objCotacao.iTipoOperacao, 0, "TipoOperacao"
    colCampoValor.Add "Cliente", objCotacao.sCliente, STRING_COTACAO_CLIENTE, "Cliente"
    colCampoValor.Add "Envio", objCotacao.iEnvio, 0, "Envio"
    colCampoValor.Add "EnvioComplemento", objCotacao.sEnvioComplemento, STRING_COTACAO_ENVIOCOMPLEMENTO, "EnvioComplemento"
    colCampoValor.Add "CodVendedor", objCotacao.iCodVendedor, 0, "CodVendedor"
    colCampoValor.Add "Indicacao", objCotacao.sIndicacao, STRING_COTACAO_INDICACAO, "Indicacao"
    colCampoValor.Add "Observacao", objCotacao.sObservacao, STRING_COTACAO_OBSERVACAO, "Observacao"
    colCampoValor.Add "DataPrevInicio", objCotacao.dtDataPrevInicio, 0, "DataPrevInicio"
    colCampoValor.Add "TipoEmbalagem", objCotacao.iTipoEmbalagem, 0, "TipoEmbalagem"
    colCampoValor.Add "Ajudantes", objCotacao.iAjudantes, 0, "Ajudantes"
    colCampoValor.Add "Carga", objCotacao.iCarga, 0, "Carga"
    colCampoValor.Add "CargaPorConta", objCotacao.iCargaPorConta, 0, "CargaPorConta"
    colCampoValor.Add "Descarga", objCotacao.iDesCarga, 0, "Descarga"
    colCampoValor.Add "DescargaPorConta", objCotacao.iDesCargaPorConta, 0, "DescargaPorConta"
    colCampoValor.Add "Ova", objCotacao.iOva, 0, "Ova"
    colCampoValor.Add "OvaPorConta", objCotacao.iOvaPorConta, 0, "OvaPorConta"
    colCampoValor.Add "Desova", objCotacao.iDesova, 0, "Desova"
    colCampoValor.Add "DesovaPorConta", objCotacao.iDesovaPorConta, 0, "DesovaPorConta"
    colCampoValor.Add "CargaSolta", objCotacao.iCargaSolta, 0, "CargaSolta"
    colCampoValor.Add "DescCargaSolta", objCotacao.sDescCargaSolta, STRING_COTACAO_DESCCARGASOLTA, "DescCargaSolta"
    colCampoValor.Add "ValorMercadoria", objCotacao.dValorMerc, 0, "ValorMercadoria"
    colCampoValor.Add "CondicaoPagto", objCotacao.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "Situacao", objCotacao.iSituacao, 0, "Situacao"
    colCampoValor.Add "Justificativa", objCotacao.iJustificativa, 0, "Justificativa"
    colCampoValor.Add "ObsResultado", objCotacao.sObsResultado, STRING_COTACAO_OBSRESULTADO, "ObsResultado"
    colCampoValor.Add "ObsDestOrigem1", sObsDestOrigem1, STRING_COTACAO_OBSDESTORIGEM, "ObsDestOrigem1"
    colCampoValor.Add "ObsDestOrigem2", sObsDestOrigem2, STRING_COTACAO_OBSDESTORIGEM, "ObsDestOrigem2"
    colCampoValor.Add "ObsDestOrigem3", sObsDestOrigem3, STRING_COTACAO_OBSDESTORIGEM, "ObsDestOrigem3"
    colCampoValor.Add "ObsDestOrigem4", sObsDestOrigem4, STRING_COTACAO_OBSDESTORIGEM, "ObsDestOrigem4"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 97250

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objCotacao As ClassCotacaoGR) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma Cotacao foi passada por parametro
    If Not (objCotacao Is Nothing) Then

            lErro = Traz_Cotacao_Tela(objCotacao)
            If lErro <> SUCESSO And lErro <> 97189 Then gError 97121
            
            If lErro = 97189 Then
                
                'limpar a tela
                Call Limpa_Tela_Cotacao
                
                'Colocar o codigo na tela
                MaskCodigo.Text = objCotacao.lCodigo
            
            End If
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr

        Case 97121

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCotacao As New ClassCotacaoGR
Dim vbMsgRes As VbMsgBoxResult
Dim lCodigo As Long

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Codigo foi preenchido
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 97159

    objCotacao.iFilialEmpresa = giFilialEmpresa
    objCotacao.lCodigo = StrParaLong(MaskCodigo.ClipText)

    'Ler a cotacao no bd
    lErro = CF("Cotacao_LeGR", objCotacao)
    If lErro <> SUCESSO And lErro <> 97170 Then gError 97218
    
    If lErro <> SUCESSO Then gError 97217

    'Confirma a exclusao da Cotacao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_COTACAO", objCotacao.lCodigo, objCotacao.iFilialEmpresa)

    'Se a resposta for sim
    If vbMsgRes = vbYes Then

        'Exclui a Cotacao
        lErro = CF("Cotacao_ExcluiGR", objCotacao)
        If lErro <> SUCESSO Then gError 97219

        'Fecha o comando das setas, se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        'Limpa a tela
        Call Limpa_Tela_Cotacao
        
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
   
    Select Case gErr

        Case 97159
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COTACAO_NAO_PREENCHIDO", gErr)

        Case 97217
            Call Rotina_Erro(vbOKOnly, "ERRO_COTACAO_NAO_CADASTRADA2", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)

        Case 97218, 97219
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a funcao que ira efetuar a gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 97167

    'limpa a tela apos a gravacao
     Call Limpa_Tela_Cotacao

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 97167

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCotacao As New ClassCotacaoGR
Dim iIndex As Integer
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Codigo da Cotacao está preenchido
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 97205

    'Verifica se a Data esta preenchida
    If Len(Trim(MaskData.ClipText)) = 0 Then gError 97206

    'Verifica se tipo de operacao esta preenchido
    If ComboTipoOperacao.ListIndex < 0 Then gError 97207

    'Verifica se a Empresa esta preenchida
    If Len(Trim(MaskCliente.ClipText)) = 0 Then gError 97212

    'verifica se tem pelo - 1 contato
    If objGridContatos.iLinhasExistentes = 0 Then gError 97213
        
    'cyntia
    For iIndice = 1 To objGridContatos.iLinhasExistentes
        If Len(Trim(GridContatos.TextMatrix(iIndice, iGrid_Contato_Col))) = 0 Then gError 99246
    Next
        
    'Verifica se o envio da cotacao foi preenchido
    If ComboEnvioCotacao.ListIndex < 0 Then gError 97208

    'verifica se tem pelo -1 cotacaodestorigem
    If objGridDestOrigem.iLinhasExistentes = 0 Then gError 97214
    
    'cyntia
    For iIndice = 1 To objGridDestOrigem.iLinhasExistentes
        If Len(Trim(GridDestOrigem.TextMatrix(iIndice, iGrid_Servico_Col))) = 0 Then gError 99259
    Next
    
    'Verifica se o tipo da embalagem foi preenchido
    If ComboTipoEmbalagem.ListIndex < 0 Then gError 97209
    
    If CheckCargaSolta.Value = vbUnchecked Then
        'verifica se tem pelo -1 cotacaocontainer
        If objGridContainer.iLinhasExistentes = 0 Then gError 97215
    End If
        
'    'Verifica se a situacao foi preenchida
'    If ComboSituacaoResultado.ListIndex < 0 Then gError 97210

    'Verifica se a data da cotacao eh maior do q a data
    'de previsao do inicio dos servicos
    If Len(Trim(MaskDataInicio.ClipText)) > 0 And Len(Trim(MaskData.ClipText)) Then
        If StrParaDate(MaskData.Text) > StrParaDate(MaskDataInicio.Text) Then gError 97380
    End If
    
'    'cyntia
'    For iIndice = 1 To objGridServicos.iLinhasExistentes
'        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Then gError 99260
'    Next
'
'    'Verifica se a quantidade de cada produto do grid serv
'    'esta preenchida
'    For iIndex = 1 To objGridServicos.iLinhasExistentes
'        If Len(Trim(GridServicos.TextMatrix(iIndex, iGrid_QuantServico_Col))) = 0 Then gError 97396
'    Next
    
    'Verifica se a quantidade de cada container do grid cont
    'esta preenchida
    For iIndex = 1 To objGridContainer.iLinhasExistentes
        If Len(Trim(GridContainer.TextMatrix(iIndex, iGrid_Quantidade_Col))) = 0 Then gError 97397
    Next
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objCotacao)
    If lErro <> SUCESSO Then gError 97211

    lErro = Trata_Alteracao(objCotacao, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98754
    
    'Cyntia
    objCotacao.sResponsavel = gsUsuario
    
    'Chama a funcao que vai efetuar, efetivamente, a gravacao
    lErro = CF("Cotacao_GravaGR", objCotacao)
    If lErro <> SUCESSO Then gError 97216
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 97205
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COTACAO_NAO_PREENCHIDO", gErr)

        Case 97206
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_COTACAO_NAO_PREENCHIDA", gErr)

        Case 97207
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOOPERACAO_COTACAO_NAO_PREENCHIDO", gErr)

        Case 97208
            Call Rotina_Erro(vbOKOnly, "ERRO_ENVIO_COTACAO_NAO_PREENCHIDO", gErr)

        Case 97209
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOEMBALAGEM_COTACAO_NAO_PREENCHIDO", gErr)
            
        Case 97210
            Call Rotina_Erro(vbOKOnly, "ERRO_SITUACAORESULTADO_NAO_PREENCHIDA", gErr)
            
        Case 97211, 97216, 98754
            
        Case 97212
            Call Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_PREENCHIDA", gErr)
        
        Case 97213
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_INEXISTENTE", gErr)
        
        Case 97214
            Call Rotina_Erro(vbOKOnly, "ERRO_COTACAOORIGEMDESTINO_INEXISTENTE", gErr)
            
        Case 97215
            Call Rotina_Erro(vbOKOnly, "ERRO_COTACAOCONTAINER_INEXISTENTE", gErr)
               
        Case 97380
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACOTACAO_MAIOR_DATAPREVINICIO", gErr)
        
        Case 97396
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDSERVICOS_NAO_PREENCHIDA", gErr, iIndex)
        
        Case 97397
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDCONTAINER_NAO_PREENCHIDA", gErr, iIndex)
        
        Case 99246
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_PREENCHIDO1", gErr, iIndice)
        
        Case 99259
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO1", gErr, iIndice)
        
        Case 99260
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO2", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function
    
End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 97161

    Call Limpa_Tela_Cotacao

    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub
    
Erro_Botaolimpar_Click:

    Select Case gErr

        Case 97161

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCod As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = Cotacao_Codigo_Automatico(lCod)
    If lErro <> SUCESSO Then gError 97132

    MaskCodigo.Text = lCod

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 97132

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoOrigemdestino_Click()

Dim objOrigemDestino As New ClassOrigemDestino
Dim colSelecao As Collection

On Error GoTo Erro_BotaoOrigemDestino_Click
    
    'Verifica se o browser está sendo chamado do controle(F3)ou pelo grid
    If (Me.ActiveControl Is Origem) Then
    
        objOrigemDestino.sOrigemDestino = Origem.Text
        
        'Chama Tela OrigemDestino
        Call Chama_Tela("OrigemDestinoLista", colSelecao, objOrigemDestino, objEventoOrigem)
    
        
    ElseIf (Me.ActiveControl Is Destino) Then
    
        objOrigemDestino.sOrigemDestino = Destino.Text
        
        'Chama Tela OrigemDestino
        Call Chama_Tela("OrigemDestinoLista", colSelecao, objOrigemDestino, objEventoDestino)
            
    Else
    'Verifica se o browser está sendo chamado pelo botão, se for
    'joga o conteudo do grid numa variável
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridServicos.Row = 0 Then gError 99261
        
        If GridServicos.Col = iGrid_OrigemServ_Col And Origem.Enabled = True Then
        
            objOrigemDestino.sOrigemDestino = GridServicos.TextMatrix(GridServicos.Row, iGrid_OrigemServ_Col)
            
            'Chama Tela OrigemDestino
            Call Chama_Tela("OrigemDestinoLista", colSelecao, objOrigemDestino, objEventoOrigem)
        ElseIf GridServicos.Col = iGrid_DestinoServ_Col And Destino.Enabled = True Then
            objOrigemDestino.sOrigemDestino = GridServicos.TextMatrix(GridServicos.Row, iGrid_DestinoServ_Col)
            
            'Chama Tela OrigemDestino
            Call Chama_Tela("OrigemDestinoLista", colSelecao, objOrigemDestino, objEventoDestino)
        Else
            gError 99258
        End If
        
    End If
    
    Exit Sub
    
Erro_BotaoOrigemDestino_Click:

    Select Case gErr
               
        Case 99258
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMDESTINO_CURSOR", gErr)
            
        Case 99261
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 99262 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoServicos_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim iIndice As Integer

On Error GoTo Erro_BotaoServicos_Click

    'Verifica se o serviço foi preenchido
    If Len(Trim(MaskProduto.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", MaskProduto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 97339

        objProduto.sCodigo = sProdutoFormatado

    End If

    If GridServicos.Row = 0 Then gError 97342
    
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico)

    'Verifica se o produto selecionado ja esta em outra linha
    For iIndice = 1 To objGridServicos.iLinhasExistentes
        If iIndice <> GridServicos.Row Then
            If GridServicos.TextMatrix(iIndice, iGrid_Produto_Col) = MaskProduto.Text Then gError 97394
        End If
    Next

    Exit Sub

Erro_BotaoServicos_Click:

    Select Case gErr

        Case 97339
        
        Case 97342
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_NAO_SELECIONADA_GRIDSERVICOS", gErr)

        Case 97394
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_JA_EXISTENTE", gErr, MaskProduto.Text, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub CheckCarga_Click()

    If CheckCarga.Value = vbChecked Then
        ComboCarga.Enabled = True
        ComboCarga.ListIndex = 0
    Else
        ComboCarga.Enabled = False
        ComboCarga.ListIndex = -1
    End If
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CheckCargaSolta_Click()

    If CheckCargaSolta.Value = vbChecked Then
        FrameTipoConteiner.Enabled = False
        Call Grid_Limpa(objGridContainer)
    Else
        FrameTipoConteiner.Enabled = True
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CheckDescarga_Click()

    If CheckDescarga.Value = vbChecked Then
        ComboDescarga.Enabled = True
        ComboDescarga.ListIndex = 0
    Else
        ComboDescarga.Enabled = False
        ComboDescarga.ListIndex = -1
    End If
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CheckDesova_Click()

    If CheckDesova.Value = vbChecked Then
        ComboDesova.Enabled = True
        ComboDesova.ListIndex = 0
    Else
        ComboDesova.Enabled = False
        ComboDesova.ListIndex = -1
    End If
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CheckOva_Click()

    If CheckOva.Value = vbChecked Then
        ComboOva.Enabled = True
        ComboOva.ListIndex = 0
    Else
        ComboOva.Enabled = False
        ComboOva.ListIndex = -1
    End If
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCarga_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCondPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCondPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboCondPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_ComboCondPagto_Validate

    'Verifica se a condpagto esta preenchida
    If Len(Trim(ComboCondPagto.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If ComboCondPagto.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(ComboCondPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 97161

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 97162
        
        If lErro = 19205 Then gError 97163

        'Coloca na Tela
        ComboCondPagto.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 97165

    Exit Sub

Erro_ComboCondPagto_Validate:

    Cancel = True

    Select Case gErr

       Case 97161, 97162

       Case 97163
            If Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo) = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 97165
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, ComboCondPagto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ComboDescarga_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboJustificativa_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboOva_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboDesova_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboEnvioCotacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboSituacaoResultado_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboTipoContainer_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboTipoEmbalagem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboTipoOperacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelVendedor_Click()

Dim colSelecao As New Collection
Dim objVendedor As New ClassVendedor
    
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub MaskAdValoren_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskCodigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigo, iAlterado)

End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validate

    If Len(Trim(MaskCodigo.ClipText)) <> 0 Then
        
        lErro = Long_Critica(MaskCodigo.ClipText)
        If lErro <> SUCESSO Then gError 97388
    
    End If
    
    Exit Sub
    
Erro_MaskCodigo_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 97388
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub MaskData_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskData_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskData, iAlterado)

End Sub

Private Sub MaskData_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskData_Validate
    
    If Len(Trim(MaskData.ClipText)) > 0 Then
        
        lErro = Data_Critica(MaskData.Text)
        If lErro <> SUCESSO Then gError 97323
        
    End If
    
    Exit Sub

Erro_MaskData_Validate:

    Cancel = True

    Select Case gErr

        Case 97323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskDataInicio_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskDataInicio_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDataInicio, iAlterado)

End Sub

Private Sub maskDataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_maskDataInicio_Validate
    
    If Len(Trim(MaskDataInicio.ClipText)) > 0 Then
        lErro = Data_Critica(MaskDataInicio.Text)
        If lErro <> SUCESSO Then gError 97324
    
    End If
    
    Exit Sub
    
Erro_maskDataInicio_Validate:
    
    Cancel = True

    Select Case gErr

        Case 97324

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub MaskPedagio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Origem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Destino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UFDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UFOrigem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskPrecoUnit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskQuantAjudantes_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskQuantAjudantes_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskQuantAjudantes, iAlterado)

End Sub

Private Sub MaskQuantAjudantes_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskQuantAjudantes_Validate

    'Pode retirar o codigo -> não, pois parece não ter
    'lógica permitir 0 na quantidade de ctr
    If Len(Trim(MaskQuantAjudantes.ClipText)) > 0 Then
        lErro = Valor_NaoNegativo_Critica(MaskQuantAjudantes.ClipText)
        If lErro <> SUCESSO Then gError 97158
    End If

    Exit Sub

Erro_MaskQuantAjudantes_Validate:

    Cancel = True

    Select Case gErr

        Case 97158

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskQuantCtr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskQuantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskValorMerc_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskValorMerc_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorMaterial As Double

On Error GoTo Erro_MaskValorMerc_Validate

    'Se Valor da Mercadoria estiver preenchido
    If Len(Trim(MaskValorMerc.ClipText)) > 0 Then
        
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(MaskValorMerc.ClipText)
        If lErro <> SUCESSO Then gError 99225
        
        'Converte o valor para double
        dValorMaterial = StrParaDbl(MaskValorMerc.Text)
        
        'formata o valor e coloca o mesmo na tela
        MaskValorMerc.Text = Format(dValorMaterial, "Standard")
    
    End If
    
    Exit Sub

Erro_MaskValorMerc_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 99225
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskVendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iVendedorAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskVendedor_Validate(Cancel As Boolean)

Dim objVendedor As New ClassVendedor
Dim lErro As Long

On Error GoTo Erro_MaskVendedor_Validate

    'Se o vendedor foi alterado
    If iVendedorAlterado = REGISTRO_ALTERADO Then
        
        'Se o vendedor esta preenchido
        If Len(Trim(MaskVendedor.ClipText)) > 0 Then
                        
            'Chama funcao TP_Vendedor_Le, que colocara na
            'Mascara, o nome reduzido do vendedor, mesmo se o
            'codigo tiver sido colocado anteriormente na mascara
            lErro = TP_Vendedor_Le(MaskVendedor, objVendedor)
            If lErro <> SUCESSO Then gError 97356
        
        End If
    
    End If

    iVendedorAlterado = 0

    Exit Sub

Erro_MaskVendedor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 97356
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim colSelecao As New Collection
Dim objCotacao As New ClassCotacaoGR

    'Verifica se o número da Cotacao foi preenchido
    If Len(Trim(MaskCodigo.ClipText)) > 0 Then objCotacao.lCodigo = StrParaLong(MaskCodigo.Text)

    'Chamada do browser
    Call Chama_Tela("CotacaoGRLista", colSelecao, objCotacao, objEventoCotacao)

End Sub

Private Sub TextContato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextDescCSolta_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextEmail_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextEnvioComplemento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextFax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextIndicacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextObservacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame atual nao corresponde ao selecionado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'torna o frame selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        
        'torna invisivel o frame atual
        Frame1(iFrameAtual).Visible = False
    
        'guarda em iframeatual o indice do novo frame
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If
    
    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TextObsResultado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextObsServico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextOrigem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextSetor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextTelefone_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    lErro = Data_Up_Down_Click(MaskData, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 97126

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 97126
            MaskData.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    lErro = Data_Up_Down_Click(MaskData, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 97127

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 97127
            MaskData.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataInicio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataInicio_DownClick

    lErro = Data_Up_Down_Click(MaskDataInicio, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 97128

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 97128
            MaskDataInicio.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataInicio_UpClick

    lErro = Data_Up_Down_Click(MaskDataInicio, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 97129

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 97129
            MaskDataInicio.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
          
    'Cyntia
    'Código do Produto
    Case MaskProduto.Name
        
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 99263

        If iProdutoPreenchido = PRODUTO_VAZIO Then
            MaskProduto.Enabled = True
        Else
            MaskProduto.Enabled = False
        End If
        
        Case MaskQuantidade.Name, MaskPrecoUnit.Name, Origem.Name, UFOrigem.Name, Destino.Name, UFDestino.Name, MaskAdValoren.Name, MaskPedagio.Name
        
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 99267

        If iProdutoPreenchido = PRODUTO_VAZIO Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If
        
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case 99263, 99267
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        If objGridInt.objGrid.Name = GridContatos.Name Then
            
            lErro = Saida_Celula_GridContatos(objGridInt)
            If lErro <> SUCESSO Then gError 97316
                   
        ElseIf objGridInt.objGrid.Name = GridServicos.Name Then
                   
            lErro = Saida_Celula_GridServicos(objGridInt)
            If lErro <> SUCESSO Then gError 97317
            
        ElseIf objGridInt.objGrid.Name = GridContainer.Name Then
        
            lErro = Saida_Celula_GridContainer(objGridInt)
            If lErro <> SUCESSO Then gError 97318
            
        ElseIf objGridInt.objGrid.Name = GridDestOrigem.Name Then
        
            lErro = Saida_Celula_GridDestOrigem(objGridInt)
            If lErro <> SUCESSO Then gError 97319
        
        End If

    End If

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97389
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 97316, 97317, 97318, 97319
        
        Case 97389
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridContatos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridContatos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Contato_Col
            
            lErro = Saida_Celula_TextContato(objGridInt)
            If lErro <> SUCESSO Then gError 97147
            
        Case iGrid_Telefone_Col
            
            lErro = Saida_Celula_TextTelefone(objGridInt)
            If lErro <> SUCESSO Then gError 97148
            
        Case iGrid_Fax_Col
        
            lErro = Saida_Celula_TextFax(objGridInt)
            If lErro <> SUCESSO Then gError 97149
            
        Case iGrid_Setor_Col
            
            lErro = Saida_Celula_TextSetor(objGridInt)
            If lErro <> SUCESSO Then gError 97150
            
        Case iGrid_Email_Col
            
            lErro = Saida_Celula_TextEmail(objGridInt)
            If lErro <> SUCESSO Then gError 97151
            
    End Select

    Saida_Celula_GridContatos = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridContatos:

    Saida_Celula_GridContatos = gErr
    
    Select Case gErr
    
        Case 97147, 97148, 97149, 97150, 97151
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_GridDestOrigem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDestOrigem

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Servico_Col
            
            lErro = Saida_Celula_textServico(objGridInt)
            If lErro <> SUCESSO Then gError 97152
            
        Case iGrid_Origem_Col
            
            lErro = Saida_Celula_textOrigem(objGridInt)
            If lErro <> SUCESSO Then gError 97153
            
        Case iGrid_Destino_Col
        
            lErro = Saida_Celula_textDestino(objGridInt)
            If lErro <> SUCESSO Then gError 97337
    
    End Select

    Saida_Celula_GridDestOrigem = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridDestOrigem:

    Saida_Celula_GridDestOrigem = gErr
    
    Select Case gErr
    
        Case 97152, 97153, 97337
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_GridContainer(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridContainer

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Container_Col

            lErro = Saida_Celula_comboTipoContainer(objGridInt)
            If lErro <> SUCESSO Then gError 97154

        Case iGrid_Quantidade_Col

            lErro = Saida_Celula_maskQuantCtr(objGridInt)
            If lErro <> SUCESSO Then gError 97155

    End Select

    Saida_Celula_GridContainer = SUCESSO

    Exit Function

Erro_Saida_Celula_GridContainer:

    Saida_Celula_GridContainer = gErr

    Select Case gErr

        Case 97154, 97155

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridServicos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridServicos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Produto_Col
            
            lErro = Saida_Celula_MaskProduto(objGridInt)
            If lErro <> SUCESSO Then gError 97156
            
        Case iGrid_QuantServico_Col
        
            lErro = Saida_Celula_MaskQuantidade(objGridInt)
            If lErro <> SUCESSO Then gError 97157
        
        Case iGrid_PrecoUnitario_Col
            
            lErro = Saida_Celula_MaskPrecoUnit(objGridInt)
            If lErro <> SUCESSO Then gError 97334
    
        Case iGrid_AdValoren_Col
            
            lErro = Saida_Celula_MaskAdValoren(objGridInt)
            If lErro <> SUCESSO Then gError 97335
        
        Case iGrid_Pedagio_Col
        
            lErro = Saida_Celula_MaskPedagio(objGridInt)
            If lErro <> SUCESSO Then gError 97336
            
        Case iGrid_OrigemServ_Col
        
            lErro = Saida_Celula_OrigemServ(objGridInt)
            If lErro <> SUCESSO Then gError 99226
                    
        Case iGrid_DestinoServ_Col
        
            lErro = Saida_Celula_DestinoServ(objGridInt)
            If lErro <> SUCESSO Then gError 99228
            
        Case iGrid_UFOrigemServ_Col
        
            lErro = Saida_Celula_UFOrigemServ(objGridInt)
            If lErro <> SUCESSO Then gError 99247
        
        Case iGrid_UFDestinoServ_Col
        
            lErro = Saida_Celula_UFDestinoServ(objGridInt)
            If lErro <> SUCESSO Then gError 99248
            
    End Select

    Saida_Celula_GridServicos = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridServicos:

    Saida_Celula_GridServicos = gErr
    
    Select Case gErr
    
        Case 97156, 97157, 97334, 97335, 97336, 99226, 99228, 99247, 99248
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Exit Function
    
End Function
Private Function Saida_Celula_UFOrigemServ(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Origem que está deixando de ser a corrente

Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_Saida_Celula_UFOrigemServ

    Set objGridInt.objControle = UFOrigem

    If Len(Trim(UFOrigem.Text)) > 0 Then
    
        lErro = TP_OrigemDestino_Le(objOrigemDestino, GridServicos.TextMatrix(GridServicos.Row, iGrid_OrigemServ_Col), UFOrigem.Text)
        If lErro <> SUCESSO Then gError 99249
            
    End If
    
    UFOrigem.Text = objOrigemDestino.sUF
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99250

    Saida_Celula_UFOrigemServ = SUCESSO

    Exit Function

Erro_Saida_Celula_UFOrigemServ:

    Saida_Celula_UFOrigemServ = gErr

    Select Case gErr

        Case 99249, 99250
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UFDestinoServ(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Destino que está deixando de ser a corrente

Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_Saida_Celula_UFDestinoServ

    Set objGridInt.objControle = UFDestino

    If Len(Trim(UFDestino.Text)) > 0 Then
    
        lErro = TP_OrigemDestino_Le(objOrigemDestino, GridServicos.TextMatrix(GridServicos.Row, iGrid_DestinoServ_Col), UFDestino.Text)
        If lErro <> SUCESSO Then gError 99251
            
    End If
    
    UFDestino.Text = objOrigemDestino.sUF
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99252

    Saida_Celula_UFDestinoServ = SUCESSO

    Exit Function

Erro_Saida_Celula_UFDestinoServ:

    Saida_Celula_UFDestinoServ = gErr

    Select Case gErr

        Case 99251, 99252
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextContato(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Contato que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextContato

    Set objGridInt.objControle = TextContato

    If Len(Trim(TextContato.Text)) > 0 Then
        If GridContatos.Row - GridContatos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97136

    Saida_Celula_TextContato = SUCESSO

    Exit Function

Erro_Saida_Celula_TextContato:

    Saida_Celula_TextContato = gErr

    Select Case gErr

        Case 97136
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextSetor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Setor que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextSetor

    Set objGridInt.objControle = TextSetor

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97137

    Saida_Celula_TextSetor = SUCESSO

    Exit Function

Erro_Saida_Celula_TextSetor:

    Saida_Celula_TextSetor = gErr

    Select Case gErr

        Case 97137
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextTelefone(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Telefone que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextTelefone

    Set objGridInt.objControle = TextTelefone

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97138

    Saida_Celula_TextTelefone = SUCESSO

    Exit Function

Erro_Saida_Celula_TextTelefone:

    Saida_Celula_TextTelefone = gErr

    Select Case gErr

        Case 97138
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextFax(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Fax que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextFax

    Set objGridInt.objControle = TextFax

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97139

    Saida_Celula_TextFax = SUCESSO

    Exit Function

Erro_Saida_Celula_TextFax:

    Saida_Celula_TextFax = gErr

    Select Case gErr

        Case 97139
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextEmail(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Email que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextEmail

    Set objGridInt.objControle = TextEmail

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97140

    Saida_Celula_TextEmail = SUCESSO

    Exit Function

Erro_Saida_Celula_TextEmail:

    Saida_Celula_TextEmail = gErr

    Select Case gErr

        Case 97140
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_textServico(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Servico que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_textServico

    Set objGridInt.objControle = TextServico

    If Len(Trim(TextServico.Text)) > 0 Then
        'Acrescenta uma linha no Grid se for o caso
        If GridDestOrigem.Row - GridDestOrigem.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97141

    Saida_Celula_textServico = SUCESSO

    Exit Function

Erro_Saida_Celula_textServico:

    Saida_Celula_textServico = gErr

    Select Case gErr

        Case 97141
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_textOrigem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Origem que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_textOrigem

    Set objGridInt.objControle = TextOrigem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97142

    Saida_Celula_textOrigem = SUCESSO

    Exit Function

Erro_Saida_Celula_textOrigem:

    Saida_Celula_textOrigem = gErr

    Select Case gErr

        Case 97142
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_textDestino(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Destino que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_textDestino

    Set objGridInt.objControle = TextDestino

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97329

    Saida_Celula_textDestino = SUCESSO

    Exit Function

Erro_Saida_Celula_textDestino:

    Saida_Celula_textDestino = gErr

    Select Case gErr

        Case 97329
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_comboTipoContainer(objGridInt As AdmGrid) As Long
'Faz a crítica da célula TipoContainer que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_comboTipoContainer

    Set objGridInt.objControle = ComboTipoContainer
    
    lErro = Verifica_Repeteco_Container
    If lErro <> SUCESSO Then gError 97355
    
    If ComboTipoContainer.ListIndex > -1 Then
        'Acrescenta uma linha no Grid se for o caso
        If GridContainer.Row - GridContainer.FixedRows = objGridInt.iLinhasExistentes Then
           objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97142

    Saida_Celula_comboTipoContainer = SUCESSO

    Exit Function

Erro_Saida_Celula_comboTipoContainer:

    Saida_Celula_comboTipoContainer = gErr

    Select Case gErr
    
        Case 97142, 97355
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_maskQuantCtr(objGridInt As AdmGrid) As Long
'Faz a crítica da célula QuantCtr que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_maskQuantCtr

    Set objGridInt.objControle = MaskQuantCtr

    If Len(Trim(MaskQuantCtr.ClipText)) > 0 Then
        lErro = Valor_Inteiro_Positivo_Critica(MaskQuantCtr.ClipText)
        If lErro <> SUCESSO Then gError 97399
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97143

    Saida_Celula_maskQuantCtr = SUCESSO

    Exit Function

Erro_Saida_Celula_maskQuantCtr:

    Saida_Celula_maskQuantCtr = gErr

    Select Case gErr

        Case 97143, 97399
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskProduto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_MaskProduto

    Set objGridInt.objControle = MaskProduto
    
    'Faz a crítica do Produto
    lErro = CF("Produto_Critica", MaskProduto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 99264
    If lErro <> SUCESSO Then gError 99265
   
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
        'Faz a crítica do produto
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO And lErro <> 97310 Then gError 97314
        
        If lErro <> SUCESSO Then
            'limpa descricao antiga
            GridServicos.TextMatrix(GridServicos.Row, iGrid_DescricaoItem_Col) = ""
            gError 97315
        End If
        
        GridServicos.TextMatrix(GridServicos.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao
        
        'Acrescenta uma linha no Grid se for o caso
        If GridServicos.Row - GridServicos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    Else

       GridServicos.TextMatrix(GridServicos.Row, iGrid_DescricaoItem_Col) = ""
              
    End If
    
    GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col) = ""
       
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97144

    Saida_Celula_MaskProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskProduto:

    Saida_Celula_MaskProduto = gErr

    Select Case gErr

        Case 97144, 97314, 99264
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 99265
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", MaskProduto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridServicos)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridServicos)
            End If

        Case 97315
            
            If (Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", MaskProduto.Text)) = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("Produto", objProduto)
            
            Else
                
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
                          
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskQuantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaskQuantidade

    Set objGridInt.objControle = MaskQuantidade

    If Len(Trim(MaskQuantidade.ClipText)) > 0 Then

        'Critica o valor informado
        lErro = Valor_Positivo_Critica(MaskQuantidade.ClipText)
        If lErro <> SUCESSO Then gError 97288
        
        
    Else
        
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Preco_Col) = ""
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97145

    'Calcula preco
    If (Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_QuantServico_Col)))) > 0 And (Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_PrecoUnitario_Col)))) > 0 Then
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Preco_Col) = Format(GridServicos.TextMatrix(GridServicos.Row, iGrid_PrecoUnitario_Col) * GridServicos.TextMatrix(GridServicos.Row, iGrid_QuantServico_Col), "standard")
    End If


    Saida_Celula_MaskQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskQuantidade:

    Saida_Celula_MaskQuantidade = gErr

    Select Case gErr

        Case 97145, 97288
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskPrecoUnit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaskPrecoUnit

    Set objGridInt.objControle = MaskPrecoUnit

    If Len(Trim(MaskPrecoUnit.ClipText)) > 0 Then

        'Critica o valor informado
        lErro = Valor_Positivo_Critica(MaskPrecoUnit.ClipText)
        If lErro <> SUCESSO Then gError 97326
        
               
    Else
        
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Preco_Col) = ""
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97325

    'Calcula preco
    If (Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_QuantServico_Col)))) > 0 And (Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_PrecoUnitario_Col)))) > 0 Then
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Preco_Col) = Format(GridServicos.TextMatrix(GridServicos.Row, iGrid_PrecoUnitario_Col) * GridServicos.TextMatrix(GridServicos.Row, iGrid_QuantServico_Col), "standard")
    End If

    Saida_Celula_MaskPrecoUnit = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskPrecoUnit:

    Saida_Celula_MaskPrecoUnit = gErr

    Select Case gErr

        Case 97325, 97326
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskAdValoren(objGridInt As AdmGrid) As Long
'Faz a crítica da célula AdValoren que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaskAdValoren

    Set objGridInt.objControle = MaskAdValoren

    If Len(Trim(MaskAdValoren.ClipText)) > 0 Then

        'Critica o valor informado
        lErro = Porcentagem_Critica(MaskAdValoren.ClipText)
        If lErro <> SUCESSO Then gError 97332
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97330

    Saida_Celula_MaskAdValoren = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskAdValoren:

    Saida_Celula_MaskAdValoren = gErr

    Select Case gErr

        Case 97330, 97332
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_OrigemServ(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Origem que está deixando de ser a corrente

Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_Saida_Celula_OrigemServ

    Set objGridInt.objControle = Origem

    If Len(Trim(Origem.Text)) > 0 Then
    
        lErro = TP_OrigemDestino_Le(objOrigemDestino, Origem.Text, GridServicos.TextMatrix(GridServicos.Row, iGrid_UFOrigemServ_Col))
        If lErro <> SUCESSO Then gError 99229
            
    End If
    
    Origem.Text = objOrigemDestino.sOrigemDestino
    GridServicos.TextMatrix(GridServicos.Row, iGrid_UFOrigemServ_Col) = objOrigemDestino.sUF

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99227

    Saida_Celula_OrigemServ = SUCESSO

    Exit Function

Erro_Saida_Celula_OrigemServ:

    Saida_Celula_OrigemServ = gErr

    Select Case gErr

        Case 99227, 99229
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DestinoServ(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Destino que está deixando de ser a corrente

Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_Saida_Celula_DestinoServ

    Set objGridInt.objControle = Destino

    If Len(Trim(Destino.Text)) > 0 Then
    
        lErro = TP_OrigemDestino_Le(objOrigemDestino, Destino.Text, GridServicos.TextMatrix(GridServicos.Row, iGrid_UFDestinoServ_Col))
        If lErro <> SUCESSO Then gError 99236
            
    End If
    
    Destino.Text = objOrigemDestino.sOrigemDestino
    GridServicos.TextMatrix(GridServicos.Row, iGrid_UFDestinoServ_Col) = objOrigemDestino.sUF

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99237

    Saida_Celula_DestinoServ = SUCESSO

    Exit Function

Erro_Saida_Celula_DestinoServ:

    Saida_Celula_DestinoServ = gErr

    Select Case gErr

        Case 99236, 99237
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskPedagio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Pedagio que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaskPedagio

    Set objGridInt.objControle = MaskPedagio

    If Len(Trim(MaskPedagio.ClipText)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(MaskPedagio.ClipText)
        If lErro <> SUCESSO Then gError 97333
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97331

    Saida_Celula_MaskPedagio = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskPedagio:

    Saida_Celula_MaskPedagio = gErr

    Select Case gErr

        Case 97331, 97333
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub GridContatos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If

End Sub

Public Sub GridContatos_GotFocus()
    Call Grid_Recebe_Foco(objGridContatos)
End Sub

Public Sub GridContatos_EnterCell()
    Call Grid_Entrada_Celula(objGridContatos, iAlterado)
End Sub

Public Sub GridContatos_LeaveCell()
    Call Saida_Celula(objGridContatos)
End Sub

Public Sub GridContatos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridContatos)
End Sub

Public Sub GridContatos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If

End Sub

Public Sub GridContatos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContatos)
End Sub

Public Sub GridContatos_RowColChange()
    
    Call Grid_RowColChange(objGridContatos)
    
End Sub

Public Sub GridContatos_Scroll()
    Call Grid_Scroll(objGridContatos)
End Sub

Public Sub GridServicos_GotFocus()
    Call Grid_Recebe_Foco(objGridServicos)
End Sub

Public Sub GridServicos_EnterCell()
    Call Grid_Entrada_Celula(objGridServicos, iAlterado)
End Sub

Public Sub GridServicos_LeaveCell()
    Call Saida_Celula(objGridServicos)
End Sub

Public Sub GridServicos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridServicos)
End Sub

Public Sub GridServicos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridServicos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServicos, iAlterado)
    End If

End Sub

Public Sub GridServicos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridServicos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServicos, iAlterado)
    End If

End Sub

Public Sub GridServicos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridServicos)
End Sub

Public Sub GridServicos_RowColChange()
    
    Call Grid_RowColChange(objGridServicos)
    
End Sub

Public Sub GridServicos_Scroll()
    Call Grid_Scroll(objGridServicos)
End Sub

Private Sub textContato_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textContato_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textContato_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextContato
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub textTelefone_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textTelefone_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textTelefone_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextTelefone
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub textFax_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textFax_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textFax_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextFax
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub textEmail_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textEmail_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textEmail_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextEmail
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub

Private Sub maskProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub maskProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = MaskProduto
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskQuantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub

Private Sub maskQuantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub maskQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = MaskQuantidade
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskPrecoUnit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub

Private Sub maskPrecoUnit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub maskPrecoUnit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = MaskPrecoUnit
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskAdValoren_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub

Private Sub Origem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub
Private Sub Destino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub

Private Sub UFOrigem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub
Private Sub UFDestino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub

Private Sub Origem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub Origem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = Origem
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Destino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub Destino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = Destino
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Sub UFOrigem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub UFOrigem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = UFOrigem
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UFDestino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub UFDestino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = UFDestino
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskAdValoren_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub maskAdValoren_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = MaskAdValoren
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskPedagio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicos)

End Sub

Private Sub maskPedagio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicos)

End Sub

Private Sub maskPedagio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicos.objControle = MaskPedagio
    lErro = Grid_Campo_Libera_Foco(objGridServicos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub textSetor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textSetor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textSetor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextSetor
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridContainer_GotFocus()
    Call Grid_Recebe_Foco(objGridContainer)
End Sub

Public Sub GridContainer_EnterCell()
    Call Grid_Entrada_Celula(objGridContainer, iAlterado)
End Sub

Public Sub GridContainer_LeaveCell()
    Call Saida_Celula(objGridContainer)
End Sub

Public Sub GridContainer_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridContainer)
End Sub

Public Sub GridContainer_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContainer, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContainer, iAlterado)
    End If

End Sub

Public Sub GridContainer_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContainer, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContainer, iAlterado)
    End If

End Sub

Public Sub GridContainer_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContainer)
End Sub

Public Sub GridContainer_RowColChange()
    
    Call Grid_RowColChange(objGridContainer)
    
End Sub

Public Sub GridContainer_Scroll()
    Call Grid_Scroll(objGridContainer)
End Sub

Public Sub GridDestOrigem_GotFocus()
    Call Grid_Recebe_Foco(objGridDestOrigem)
End Sub

Public Sub GridDestOrigem_EnterCell()
    Call Grid_Entrada_Celula(objGridDestOrigem, iAlterado)
End Sub

Public Sub GridDestOrigem_LeaveCell()
    Call Saida_Celula(objGridDestOrigem)
End Sub

Public Sub GridDestOrigem_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridDestOrigem)
End Sub

Public Sub GridDestOrigem_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDestOrigem, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDestOrigem, iAlterado)
    End If

End Sub

Public Sub GridDestOrigem_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDestOrigem, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDestOrigem, iAlterado)
    End If

End Sub

Public Sub GridDestOrigem_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridDestOrigem)
End Sub

Public Sub GridDestOrigem_RowColChange()
    
    Call Grid_RowColChange(objGridDestOrigem)
    
End Sub

Public Sub GridDestOrigem_Scroll()
    Call Grid_Scroll(objGridDestOrigem)
End Sub

Private Sub textServico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDestOrigem)

End Sub

Private Sub textServico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDestOrigem)

End Sub

Private Sub textServico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDestOrigem.objControle = TextServico
    lErro = Grid_Campo_Libera_Foco(objGridDestOrigem)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub textOrigem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDestOrigem)

End Sub

Private Sub textOrigem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDestOrigem)

End Sub

Private Sub textOrigem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDestOrigem.objControle = TextOrigem
    lErro = Grid_Campo_Libera_Foco(objGridDestOrigem)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub textDestino_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDestOrigem)

End Sub

Private Sub textDestino_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDestOrigem)

End Sub

Private Sub textDestino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDestOrigem.objControle = TextDestino
    lErro = Grid_Campo_Libera_Foco(objGridDestOrigem)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComboTipoContainer_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContainer)

End Sub

Private Sub ComboTipoContainer_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContainer)

End Sub

Private Sub ComboTipoContainer_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContainer.objControle = ComboTipoContainer
    lErro = Grid_Campo_Libera_Foco(objGridContainer)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskQuantCtr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContainer)

End Sub

Private Sub maskQuantCtr_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContainer)

End Sub

Private Sub maskQuantCtr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContainer.objControle = MaskQuantCtr
    lErro = Grid_Campo_Libera_Foco(objGridContainer)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Verifica_Repeteco_Container() As Long
'Verifica se ja existe o container em questao em alguma linha

Dim iIndice As Integer

On Error GoTo Erro_Verifica_Repeteco_Container

    For iIndice = 1 To objGridContainer.iLinhasExistentes
        If iIndice <> GridContainer.Row Then
            If GridContainer.TextMatrix(iIndice, iGrid_Container_Col) = ComboTipoContainer.Text Then gError 97354
        End If
    Next

    Exit Function

Erro_Verifica_Repeteco_Container:

    Verifica_Repeteco_Container = gErr

    Select Case gErr

        Case 97354
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAINER_JA_EXISTENTE", gErr, ComboTipoContainer.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Carrega_CondicaoPagamento() As Long
'Funcao Importada de GlobaisTelasFat/CTPedidoVenda

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 97390

    For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na List da Combo CondicaoPagamento
        ComboCondPagto.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        ComboCondPagto.ItemData(ComboCondPagto.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case 97390

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

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

    'Libera o comando de setas
    lErro = ComandoSeta_Liberar(Me.Name)
    Set objEventoOrigem = Nothing
    Set objEventoDestino = Nothing
    Set objEventoCotacao = Nothing
    Set objEventoServico = Nothing
    Set objEventoVendedor = Nothing
    
End Sub

Private Function Cotacao_Codigo_Automatico(lCod As Long) As Long
'funcao que gera o codigo automatico

Dim lErro As Long

On Error GoTo Erro_Cotacao_Codigo_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_COTACAO", "CotacaoGR", "Codigo", lCod)
    If lErro <> SUCESSO Then gError 97133

    Cotacao_Codigo_Automatico = SUCESSO

    Exit Function

Erro_Cotacao_Codigo_Automatico:

    Cotacao_Codigo_Automatico = gErr

    Select Case gErr

        Case 97133
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub objEventoCotacao_evSelecao(obj1 As Object)

Dim objCotacao As ClassCotacaoGR
Dim lErro As Long

On Error GoTo Erro_objEventoCotacao_evSelecao

    Set objCotacao = obj1

    lErro = Traz_Cotacao_Tela(objCotacao)
    If lErro <> SUCESSO And lErro <> 97189 Then gError 97134
   
    If lErro <> SUCESSO Then gError 97135
   
    'Fecha o comando das setas, se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCotacao_evSelecao:

    Select Case gErr

        Case 97134
        
        Case 97135
            Call Rotina_Erro(vbOKOnly, "ERRO_COTACAO_NAO_CADASTRADA2", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function Produto_Saida_Celula(Optional objProduto As ClassProduto) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Produto_Saida_Celula

    'Critica o Produto
    If objProduto Is Nothing Then
        Set objProduto = New ClassProduto
        lErro = CF("Produto_Critica", MaskProduto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 42186
        If lErro <> SUCESSO Then gError 42191
    End If

    'Verifica se é de Faturamento
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 42193
    
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then gError 42190

    MaskProduto.PromptInclude = False
    MaskProduto.Text = sProdutoEnxuto
    MaskProduto.PromptInclude = True

    'Verifica se está no Grid
    For iIndice = 1 To objGridServicos.iLinhasExistentes
        If iIndice <> GridServicos.Row Then If GridServicos.TextMatrix(iIndice, iGrid_Produto_Col) = MaskProduto.Text Then gError 42192
    Next

    
    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case 42186, 42915

        Case 42190
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 42191
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", MaskProduto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridServicos)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridServicos)
            End If


        Case 42192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, MaskProduto.Text, iIndice)

        Case 42193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim iIndice As Integer
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoServico_evSelecao
    
    'verifica se tem alguma linha do Grid selecionada
    If GridServicos.Row = 0 Then gError 97340

    'Verifica se o Produto está preenchido
    If Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col))) = 0 Then

        Set objProduto = obj1
        
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 97341

        'Traz para a tela o servico e a descricao
        MaskProduto.PromptInclude = False
        MaskProduto.Text = sProdutoEnxuto
        MaskProduto.PromptInclude = True
            
        'Coloca o produto
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col) = MaskProduto.Text
                
        lErro = Produto_Saida_Celula(objProduto)
        If lErro <> SUCESSO Then gError 99266

        'Verifica se o browser está sendo chamado pelo botão, se for
        'joga no grid a descrição e o produto
        If Not Me.ActiveControl Is MaskProduto Then
            GridServicos.TextMatrix(GridServicos.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao
            If GridServicos.Row - GridServicos.FixedRows = objGridServicos.iLinhasExistentes Then
                objGridServicos.iLinhasExistentes = objGridServicos.iLinhasExistentes + 1
            End If
        Else
            GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col) = ""
        End If
        
        Me.Show

    End If

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr
        
        Case 97340
        
        Case 99266
            GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col) = ""
            MaskProduto.PromptInclude = False
            MaskProduto.Text = ""
            MaskProduto.PromptInclude = True
            
        Case 97341
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim lErro As Long

    Set objVendedor = obj1

    MaskVendedor.Text = objVendedor.sNomeReduzido

    Me.Show

End Sub
Private Sub objEventoDestino_evSelecao(obj1 As Object)

Dim objOrigemDestino As ClassOrigemDestino
Dim lErro As Long

On Error GoTo Erro_objEventoDestino_evSelecao

    Set objOrigemDestino = obj1

    'Move Destino e UF para o grid de serviços
    Destino.Text = objOrigemDestino.sOrigemDestino
    GridServicos.TextMatrix(GridServicos.Row, iGrid_UFDestinoServ_Col) = objOrigemDestino.sUF
    
    Me.Show

    Exit Sub
    
Erro_objEventoDestino_evSelecao:

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoOrigem_evSelecao(obj1 As Object)

Dim objOrigemDestino As ClassOrigemDestino
Dim lErro As Long

On Error GoTo Erro_objEventoOrigem_evSelecao

    Set objOrigemDestino = obj1

    'Move Origem e UF para o grid de serviços
    Origem.Text = objOrigemDestino.sOrigemDestino
    GridServicos.TextMatrix(GridServicos.Row, iGrid_UFOrigemServ_Col) = objOrigemDestino.sUF
    
    Me.Show

    Exit Sub
    
Erro_objEventoOrigem_evSelecao:

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Sub Limpa_Tela_Cotacao()

Dim iIndice As Integer
Dim lErro As Long
Dim objUsuarios As New ClassUsuarios
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Limpa_Tela_Cotacao

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
        
        'Combos
        'ComboTipoOperacao.Text = ""
        ComboTipoOperacao.ListIndex = -1
        'ComboEnvioCotacao.Text = ""
        ComboEnvioCotacao.ListIndex = -1
        'ComboTipoEmbalagem.Text = ""
        'ComboTipoEmbalagem.ListIndex = -1
        ComboCondPagto.Text = ""
        ComboCondPagto.ListIndex = -1
        'ComboSituacaoResultado.Text = ""
        'ComboSituacaoResultado.ListIndex = 0
        ComboSituacaoResultado.ListIndex = -1
        'ComboJustificativa.Text = ""
        ComboJustificativa.ListIndex = -1
        ComboOva.Enabled = False
        ComboDesova.Enabled = False
        ComboCarga.Enabled = False
        ComboDescarga.Enabled = False
        
        'Grids
        Call Grid_Limpa(objGridContatos)
        Call Grid_Limpa(objGridContainer)
        Call Grid_Limpa(objGridDestOrigem)
        'Call Grid_Limpa(objGridServicos)
        
        'Checks
        CheckCarga.Value = vbUnchecked
        CheckOva.Value = vbUnchecked
        CheckDesova.Value = vbUnchecked
        CheckDescarga.Value = vbUnchecked
        CheckCargaSolta.Value = vbUnchecked
    
        'Outros
        MaskData.PromptInclude = False
        MaskData.Text = Format(gdtDataHoje, "dd/mm/yy")
        MaskData.PromptInclude = True
            
     Exit Sub

Erro_Limpa_Tela_Cotacao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objCotacao As ClassCotacaoGR) As Long

Dim objVendedor As New ClassVendedor
Dim sAux As String
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objCotacao.dtData = StrParaDate(MaskData.Text)
    objCotacao.dtDataPrevInicio = StrParaDate(MaskDataInicio.Text)
    
    If Len(Trim(MaskQuantAjudantes.ClipText)) >= 0 Then objCotacao.iAjudantes = StrParaLong(MaskQuantAjudantes.Text)
    
    If CheckCarga.Value = vbUnchecked Then
        objCotacao.iCarga = 0
    Else
        objCotacao.iCarga = 1
        objCotacao.iCargaPorConta = ComboCarga.ItemData(ComboCarga.ListIndex)
    End If
    
    If CheckDescarga.Value = vbUnchecked Then
        objCotacao.iDesCarga = 0
    Else
        objCotacao.iDesCarga = 1
        objCotacao.iDesCargaPorConta = ComboDescarga.ItemData(ComboDescarga.ListIndex)
    End If
    
    If CheckOva.Value = vbUnchecked Then
        objCotacao.iOva = 0
    Else
        objCotacao.iOva = 1
        objCotacao.iOvaPorConta = ComboOva.ItemData(ComboOva.ListIndex)
    End If
    
    If CheckDesova.Value = vbUnchecked Then
        objCotacao.iDesova = 0
    Else
        objCotacao.iDesova = 1
        objCotacao.iDesovaPorConta = ComboDesova.ItemData(ComboDesova.ListIndex)
    End If
      
    If CheckCargaSolta.Value = vbUnchecked Then
        objCotacao.iCargaSolta = 0
    Else
        objCotacao.iCargaSolta = 1
    End If
    
    objCotacao.sDescCargaSolta = TextDescCSolta.Text
    If MaskVendedor.ClipText <> "" Then
       
       objVendedor.sNomeReduzido = MaskVendedor.ClipText
       
       'para pegar o codigo do vendedor...
       'O caso de nao encontrar (erro: 25008) nao eh
       'tratado propositalmente.. esse erro provavelmente nunca sera
       'disparado, mas caso aconteca, o codigo do vendedor sera 0 devido
       'a inicializacao do obj
       lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
       If lErro <> SUCESSO And lErro <> 25008 Then gError 97391
       
    End If
    
    objCotacao.dValorMerc = StrParaDbl(MaskValorMerc.Text)
    
    'se achar o vendedor, guarda o codigo, senao, o valor eh 0 (inicializacao do obj)
    objCotacao.iCodVendedor = objVendedor.iCodigo
        
    If ComboCondPagto.ListIndex >= 0 Then
        objCotacao.iCondicaoPagto = Codigo_Extrai(ComboCondPagto.Text)
    End If
  
    If ComboEnvioCotacao.ListIndex >= 0 Then
        objCotacao.iEnvio = ComboEnvioCotacao.ItemData(ComboEnvioCotacao.ListIndex)
    End If
    
    objCotacao.iFilialEmpresa = giFilialEmpresa
    
    'no futuro?? objCotacao.iJustificativa = ComboJustificativa.ItemData(ComboJustificativa.ListIndex)

    If ComboSituacaoResultado.ListIndex >= 0 Then
        objCotacao.iSituacao = ComboSituacaoResultado.ItemData(ComboSituacaoResultado.ListIndex)
    End If
    
    If ComboTipoEmbalagem.ListIndex >= 0 Then
        objCotacao.iTipoEmbalagem = ComboTipoEmbalagem.ItemData(ComboTipoEmbalagem.ListIndex)
    End If
    
    If ComboTipoOperacao.ListIndex >= 0 Then
        objCotacao.iTipoOperacao = ComboTipoOperacao.ItemData(ComboTipoOperacao.ListIndex)
    End If
    
    objCotacao.lCodigo = StrParaLong(MaskCodigo.ClipText)
    
    objCotacao.sCliente = MaskCliente.ClipText
    
    objCotacao.sEnvioComplemento = TextEnvioComplemento.Text
    
    objCotacao.sIndicacao = TextIndicacao.Text
    
    objCotacao.sObsDestOrigem = TextObsServico.Text
    
    objCotacao.sObservacao = TextObservacao.Text
    
    objCotacao.sObsResultado = TextObsResultado.Text
    
    lErro = Move_Grids_Memoria(objCotacao)
    If lErro <> SUCESSO Then gError 97358

    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 97358
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Function Move_Grids_Memoria(objCotacao As ClassCotacaoGR) As Long
'Move itens dos grids para o objCotacao

Dim objCotacaoServico As ClassCotacaoServico
Dim objCotacaoContainer As ClassCotacaoContainer
Dim objCotacaoOrigemDestino As ClassCotacaoOrigemDestino
Dim objContato As ClassContato
Dim sProduto As String
Dim iLinha As Integer
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_Move_Grids_Memoria

    For iLinha = 1 To objGridContatos.iLinhasExistentes
        
        Set objContato = New ClassContato
        
        'Carrega os dados em objContato
        objContato.sContato = GridContatos.TextMatrix(iLinha, iGrid_Contato_Col)
        objContato.sEmail = GridContatos.TextMatrix(iLinha, iGrid_Email_Col)
        objContato.sFax = GridContatos.TextMatrix(iLinha, iGrid_Fax_Col)
        objContato.sSetor = GridContatos.TextMatrix(iLinha, iGrid_Setor_Col)
        objContato.sTelefone = GridContatos.TextMatrix(iLinha, iGrid_Telefone_Col)

        objCotacao.colContato.Add objContato
        
    Next

    For iLinha = 1 To objGridServicos.iLinhasExistentes

        Set objCotacaoServico = New ClassCotacaoServico
    
        'Carrega os dados em objCotacaoServico
        objCotacaoServico.dAdValoren = PercentParaDbl(GridServicos.TextMatrix(iLinha, iGrid_AdValoren_Col))
        objCotacaoServico.dPedagio = StrParaDbl(GridServicos.TextMatrix(iLinha, iGrid_Pedagio_Col))
        objCotacaoServico.dPrecoUnitario = StrParaDbl(GridServicos.TextMatrix(iLinha, iGrid_PrecoUnitario_Col))
        objCotacaoServico.dQuantidade = StrParaDbl(GridServicos.TextMatrix(iLinha, iGrid_QuantServico_Col))
        
        If Len(GridServicos.TextMatrix(iLinha, iGrid_OrigemServ_Col)) <> 0 Then
                        
            objOrigemDestino.sOrigemDestino = GridServicos.TextMatrix(iLinha, iGrid_OrigemServ_Col)
            objOrigemDestino.sUF = GridServicos.TextMatrix(iLinha, iGrid_UFOrigemServ_Col)
            
            lErro = CF("OrigemDestino_Le_NomeUF", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96860 Then gError 99238
            
            If lErro = 96860 Then gError 99239
            
            objCotacaoServico.iOrigem = objOrigemDestino.iCodigo
            
        End If
        
        If Len(GridServicos.TextMatrix(iLinha, iGrid_DestinoServ_Col)) <> 0 Then
            
            objOrigemDestino.sOrigemDestino = GridServicos.TextMatrix(iLinha, iGrid_DestinoServ_Col)
            objOrigemDestino.sUF = GridServicos.TextMatrix(iLinha, iGrid_UFDestinoServ_Col)
            
            lErro = CF("OrigemDestino_Le_NomeUF", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96860 Then gError 99240
            
            If lErro = 96860 Then gError 99241
            
            objCotacaoServico.iDestino = objOrigemDestino.iCodigo
            
        End If
        
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 97392
                
        objCotacaoServico.sProduto = sProduto
                
        objCotacao.colCotacaoServico.Add objCotacaoServico
    
    Next
    
    For iLinha = 1 To objGridDestOrigem.iLinhasExistentes

        Set objCotacaoOrigemDestino = New ClassCotacaoOrigemDestino

        'Carrega os dados em objGridDestOrigem
        objCotacaoOrigemDestino.sServico = GridDestOrigem.TextMatrix(iLinha, iGrid_Servico_Col)
        objCotacaoOrigemDestino.sOrigem = GridDestOrigem.TextMatrix(iLinha, iGrid_Origem_Col)
        objCotacaoOrigemDestino.sDestino = GridDestOrigem.TextMatrix(iLinha, iGrid_Destino_Col)
    
        objCotacao.colCotacaoOrigemDestino.Add objCotacaoOrigemDestino
    
    Next

    If CheckCargaSolta.Value = vbUnchecked Then

        For iLinha = 1 To objGridContainer.iLinhasExistentes

            Set objCotacaoContainer = New ClassCotacaoContainer

            'Carrega os dados em objGridCotacaoContainer
            objCotacaoContainer.iQuantidade = StrParaInt(GridContainer.TextMatrix(iLinha, iGrid_Quantidade_Col))
            objCotacaoContainer.iTipoContainer = StrParaInt(Codigo_Extrai(GridContainer.TextMatrix(iLinha, iGrid_Container_Col)))

            objCotacao.colCotacaoContainer.Add objCotacaoContainer

        Next


    End If

    Move_Grids_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Grids_Memoria:

    Move_Grids_Memoria = gErr
    
    Select Case gErr
    
        Case 97392, 99238, 99240
        
        Case 99239
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_EXISTENTE", gErr, objOrigemDestino.iCodigo)
        
        Case 99241
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_EXISTENTE", gErr, objOrigemDestino.iCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long
'Formata o produto deixando-o pronto para ser trazido para a tela

Dim lErro As Long
Dim iServicoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sServico As String

On Error GoTo Erro_Traz_Produto_Tela

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", MaskProduto.Text, objProduto, iServicoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 97309
    
    If lErro = 51381 Then gError 97310

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sServico)
    If lErro <> SUCESSO Then gError 97311
       
    MaskProduto.PromptInclude = False
    MaskProduto.Text = sServico
    MaskProduto.PromptInclude = True
    
    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objGridServicos.iLinhasExistentes
        If iIndice <> GridServicos.Row Then
            If GridServicos.TextMatrix(iIndice, iGrid_Produto_Col) = MaskProduto.Text Then gError 97313
        End If
    Next

    'Verifica se é de Faturamento
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 97312
    
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 97309, 97310

        Case 97311
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sNomeReduzido)
        
        Case 97312
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PODE_SER_VENDIDO", gErr, objProduto.sCodigo)

        Case 97313
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_JA_EXISTENTE", gErr, MaskProduto.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

'Caso o usuario queira acessar o browser através da tecla F3.
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is MaskProduto Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is MaskCodigo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is MaskVendedor Then
            Call LabelVendedor_Click
        ElseIf Me.ActiveControl Is Origem Then
            Call BotaoOrigemdestino_Click
        ElseIf Me.ActiveControl Is Destino Then
            Call BotaoOrigemdestino_Click
        End If
        
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
    
        Call BotaoProxNum_Click
    
    End If

End Sub

'??? perguntar se deve subir
Function Valor_Inteiro_Positivo_Critica(sValor As String) As Long
'critica se o valor passado como parametro é
'positivo e inteiro. Se estiver tudo ok retorna SUCESSO
'Essa funcao deve ser usada para critica um campo no qual
'a máscara só permita que sejam digitados os caracteres
'numéricos: 1, 2, 3, 4, 5, 6, 7, 8, 9, 0

Dim iTeste As Integer
Dim lErro As Long

On Error GoTo Erro_Valor_Inteiro_Positivo_Critica

    iTeste = CInt(sValor)
    
    If iTeste = 0 Then gError 97398
    
    Valor_Inteiro_Positivo_Critica = SUCESSO

    Exit Function

Erro_Valor_Inteiro_Positivo_Critica:

    Valor_Inteiro_Positivo_Critica = gErr
    
    Select Case gErr
    
        Case 97398
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_POSITIVO", gErr, sValor)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, sValor)
            
    End Select
        
    Exit Function

End Function



''**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Pedido de Cotação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Cotacao"

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



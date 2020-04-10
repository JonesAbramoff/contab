VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CompServico 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   2
      Left            =   150
      TabIndex        =   52
      Top             =   1005
      Width           =   9270
      Begin VB.Frame Frame2 
         Caption         =   "Container"
         Height          =   1200
         Left            =   30
         TabIndex        =   78
         Top             =   3135
         Width           =   9225
         Begin VB.TextBox TextLacreContainer 
            Height          =   315
            Left            =   5205
            MaxLength       =   20
            TabIndex        =   13
            Top             =   825
            Width           =   2145
         End
         Begin MSMask.MaskEdBox MaskCodigoContainer 
            Height          =   315
            Left            =   5205
            TabIndex        =   10
            Top             =   300
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "????-###.###-#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskTaraContainer 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   825
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskValorContainer 
            Height          =   315
            Left            =   7890
            TabIndex        =   11
            Top             =   285
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   2
            Left            =   7275
            TabIndex        =   104
            Top             =   360
            Width           =   510
         End
         Begin VB.Label LabelTipoContainer 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   1320
            TabIndex        =   99
            Top             =   330
            Width           =   2505
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Lacre:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   21
            Left            =   4530
            TabIndex        =   82
            Top             =   885
            Width           =   555
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   795
            TabIndex        =   81
            Top             =   375
            Width           =   450
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   6
            Left            =   4365
            TabIndex        =   80
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tara:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   7
            Left            =   795
            TabIndex        =   79
            Top             =   885
            Width           =   465
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados do Navio"
         Height          =   3075
         Left            =   30
         TabIndex        =   58
         Top             =   45
         Width           =   9225
         Begin VB.Frame Frame3 
            Caption         =   "Chegada"
            Height          =   735
            Left            =   195
            TabIndex        =   64
            Top             =   2250
            Width           =   4356
            Begin VB.Label LabelHoraChegada 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3030
               TabIndex        =   68
               Top             =   375
               Width           =   855
            End
            Begin VB.Label LabelDataChegada 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   945
               TabIndex        =   67
               Top             =   375
               Width           =   1095
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
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
               Index           =   11
               Left            =   2475
               TabIndex        =   66
               Top             =   420
               Width           =   465
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   12
               Left            =   375
               TabIndex        =   65
               Top             =   390
               Width           =   480
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "DeadLine"
            Height          =   750
            Left            =   4800
            TabIndex        =   59
            Top             =   2235
            Width           =   4356
            Begin VB.Label LabelHoraDeadLine 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3108
               TabIndex        =   63
               Top             =   372
               Width           =   852
            End
            Begin VB.Label LabelDataDeadLine 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   960
               TabIndex        =   62
               Top             =   372
               Width           =   1092
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
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
               Height          =   192
               Index           =   14
               Left            =   2544
               TabIndex        =   61
               Top             =   426
               Width           =   480
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H80000008&
               Height          =   192
               Index           =   13
               Left            =   408
               TabIndex        =   60
               Top             =   426
               Width           =   480
            End
         End
         Begin MSComCtl2.UpDown UpDownDemurrage 
            Height          =   330
            Left            =   6345
            TabIndex        =   83
            Top             =   345
            Width           =   240
            _ExtentX        =   344
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox MaskDemurrage 
            Height          =   315
            Left            =   5385
            TabIndex        =   9
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label labelPorto 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5385
            TabIndex        =   118
            Top             =   1755
            Width           =   3765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Porto Desemb.:"
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
            Index           =   28
            Left            =   4035
            TabIndex        =   117
            Top             =   1845
            Width           =   1320
         End
         Begin VB.Label labelBooking 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1200
            TabIndex        =   116
            Top             =   1740
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Booking:"
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
            Index           =   27
            Left            =   375
            TabIndex        =   115
            Top             =   1800
            Width           =   765
         End
         Begin VB.Label LabelProgNavio 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   1200
            TabIndex        =   100
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Demurrage:"
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
            Index           =   22
            Left            =   4350
            TabIndex        =   84
            Top             =   405
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Armador:"
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
            Index           =   15
            Left            =   330
            TabIndex        =   77
            Top             =   870
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Viagem:"
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
            Index           =   16
            Left            =   4650
            TabIndex        =   76
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Navio:"
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
            Index           =   17
            Left            =   525
            TabIndex        =   75
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Id. Viagem:"
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
            Left            =   120
            TabIndex        =   74
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ag. Marítimo:"
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
            Index           =   18
            Left            =   4185
            TabIndex        =   73
            Top             =   870
            Width           =   1110
         End
         Begin VB.Label LabelArmador 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1200
            TabIndex        =   72
            Top             =   795
            Width           =   2640
         End
         Begin VB.Label LabelAgMaritimo 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5385
            TabIndex        =   71
            Top             =   810
            Width           =   3765
         End
         Begin VB.Label LabelNavio 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1200
            TabIndex        =   70
            Top             =   1275
            Width           =   2655
         End
         Begin VB.Label LabelViagem 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5385
            TabIndex        =   69
            Top             =   1275
            Width           =   3765
         End
      End
      Begin VB.TextBox TextObs 
         Height          =   315
         Left            =   1380
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   4350
         Width           =   7815
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   54
         Top             =   4410
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   1
      Left            =   165
      TabIndex        =   41
      Top             =   1020
      Width           =   9270
      Begin VB.TextBox TextUM 
         Height          =   315
         Left            =   7500
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1575
         Width           =   915
      End
      Begin MSComCtl2.UpDown UpDownDataEmissao 
         Height          =   330
         Left            =   8445
         TabIndex        =   36
         Top             =   135
         Width           =   240
         _ExtentX        =   344
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskDataEmissao 
         Height          =   315
         Left            =   7485
         TabIndex        =   3
         Top             =   150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskQuantMaterial 
         Height          =   315
         Left            =   4830
         TabIndex        =   6
         Top             =   1575
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskValorMerc 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   2095
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskQuantidade 
         Height          =   315
         Left            =   7470
         TabIndex        =   5
         Top             =   660
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskServico 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelPedagio 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   7500
         TabIndex        =   114
         Top             =   2595
         Width           =   1170
      End
      Begin VB.Label LabelAdValoren 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4830
         TabIndex        =   112
         Top             =   2595
         Width           =   1155
      End
      Begin VB.Label LabelPrecoUnitario 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1440
         TabIndex        =   111
         Top             =   2600
         Width           =   1755
      End
      Begin VB.Label labelDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1440
         TabIndex        =   102
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   420
         TabIndex        =   101
         Top             =   1140
         Width           =   930
      End
      Begin VB.Label LabelDespachante 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1440
         TabIndex        =   98
         Top             =   4260
         Width           =   4545
      End
      Begin VB.Label LabelMaterial 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1440
         TabIndex        =   97
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label LabelTipoEmbalagem 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4830
         TabIndex        =   96
         Top             =   2070
         Width           =   2205
      End
      Begin VB.Label LabelCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1440
         TabIndex        =   95
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
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
         Index           =   26
         Left            =   7095
         TabIndex        =   94
         Top             =   3780
         Width           =   315
      End
      Begin VB.Label LabelUFDestino 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   7500
         TabIndex        =   93
         Top             =   3705
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
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
         Index           =   25
         Left            =   7095
         TabIndex        =   92
         Top             =   3210
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   705
         TabIndex        =   91
         Top             =   3210
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pedágio:"
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
         Left            =   6645
         TabIndex        =   90
         Top             =   2675
         Width           =   765
      End
      Begin VB.Label LabelUFOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   7500
         TabIndex        =   89
         Top             =   3135
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
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
         Index           =   24
         Left            =   645
         TabIndex        =   88
         Top             =   3780
         Width           =   720
      End
      Begin VB.Label LabelDestino 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1440
         TabIndex        =   87
         Top             =   3705
         Width           =   4545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Preço Unitário:"
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
         Index           =   9
         Left            =   75
         TabIndex        =   86
         Top             =   2675
         Width           =   1290
      End
      Begin VB.Label LabelOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1440
         TabIndex        =   85
         Top             =   3135
         Width           =   4545
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6360
         TabIndex        =   56
         Top             =   690
         Width           =   1050
      End
      Begin VB.Label label8 
         AutoSize        =   -1  'True
         Caption         =   "Despachante:"
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
         Index           =   19
         Left            =   180
         TabIndex        =   55
         Top             =   4335
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AD - Valoren %:"
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
         Index           =   8
         Left            =   3405
         TabIndex        =   51
         Top             =   2675
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Embalagem:"
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
         Left            =   3285
         TabIndex        =   50
         Top             =   2145
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Mercad.:"
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
         Left            =   105
         TabIndex        =   49
         Top             =   2145
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   6930
         TabIndex        =   48
         Top             =   1635
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Quant. Material:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   3375
         TabIndex        =   47
         Top             =   1635
         Width           =   1380
      End
      Begin VB.Label label6 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
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
         Left            =   615
         TabIndex        =   46
         Top             =   1635
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   6180
         TabIndex        =   45
         Top             =   195
         Width           =   1230
      End
      Begin VB.Label labelServico 
         AutoSize        =   -1  'True
         Caption         =   "Serviço:"
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
         Left            =   645
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   44
         Top             =   690
         Width           =   720
      End
      Begin VB.Label ClienteLabel 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   705
         TabIndex        =   43
         Top             =   195
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4695
      Index           =   3
      Left            =   135
      TabIndex        =   53
      Top             =   1020
      Width           =   9270
      Begin VB.CommandButton BotaoDocumentos 
         Caption         =   "Documentos"
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
         Left            =   4620
         TabIndex        =   103
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton BotaoImprimirComprovante 
         Caption         =   "Imprimir Comprovante"
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
         Left            =   6900
         TabIndex        =   34
         Top             =   4200
         Width           =   2205
      End
      Begin VB.TextBox TextObservacao 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4848
         MaxLength       =   250
         TabIndex        =   33
         Top             =   2100
         Width           =   2105
      End
      Begin MSMask.MaskEdBox MaskMotorista 
         Height          =   228
         Left            =   3444
         TabIndex        =   32
         Top             =   2100
         Width           =   1332
         _ExtentX        =   2328
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskPlacaCarreta 
         Height          =   225
         Left            =   1980
         TabIndex        =   31
         Top             =   2100
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskPlacaCaminhao 
         Height          =   225
         Left            =   435
         TabIndex        =   30
         Top             =   2130
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDocExtHoraRec 
         Height          =   225
         Left            =   6330
         TabIndex        =   29
         Top             =   1665
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDocExtDataRec 
         Height          =   225
         Left            =   5100
         TabIndex        =   28
         Top             =   1650
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDocExtNumero 
         Height          =   225
         Left            =   2730
         TabIndex        =   26
         Top             =   1710
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDocExtTipo 
         Height          =   225
         Left            =   1695
         TabIndex        =   25
         Top             =   1710
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDocExtDataEmi 
         Height          =   225
         Left            =   3810
         TabIndex        =   27
         Top             =   1650
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDocIntDataEmi 
         Height          =   225
         Left            =   420
         TabIndex        =   24
         Top             =   1695
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDocIntNumero 
         Height          =   225
         Left            =   6585
         TabIndex        =   23
         Top             =   1350
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDocIntTipo 
         Height          =   225
         Left            =   5460
         TabIndex        =   22
         Top             =   1290
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDataInicio 
         Height          =   225
         Left            =   480
         TabIndex        =   18
         Top             =   1290
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskHoraInicio 
         Height          =   225
         Left            =   1650
         TabIndex        =   19
         Top             =   1305
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskHoraFim 
         Height          =   225
         Left            =   4140
         TabIndex        =   21
         Top             =   1305
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskHoraPrev 
         Height          =   225
         Left            =   6600
         TabIndex        =   17
         Top             =   990
         Width           =   1296
         _ExtentX        =   2275
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDataFim 
         Height          =   225
         Left            =   2880
         TabIndex        =   20
         Top             =   1260
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDataPrev 
         Height          =   225
         Left            =   5430
         TabIndex        =   16
         Top             =   960
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.TextBox TextDescItemServico 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2355
         MaxLength       =   50
         TabIndex        =   38
         Top             =   870
         Width           =   5340
      End
      Begin MSMask.MaskEdBox MaskItemServico 
         Height          =   225
         Left            =   930
         TabIndex        =   37
         Top             =   975
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItemServico 
         Height          =   3945
         Left            =   45
         TabIndex        =   15
         Top             =   330
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   6959
         _Version        =   393216
         Rows            =   16
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6840
      ScaleHeight     =   495
      ScaleWidth      =   2565
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   90
      Width           =   2625
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2070
         Picture         =   "CompServicoGR.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1590
         Picture         =   "CompServicoGR.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1065
         Picture         =   "CompServicoGR.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   555
         Picture         =   "CompServicoGR.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         Picture         =   "CompServicoGR.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboStatus 
      Height          =   315
      ItemData        =   "CompServicoGR.ctx":0A96
      Left            =   4950
      List            =   "CompServicoGR.ctx":0AA3
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   195
      Width           =   1725
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   312
      Left            =   3885
      Picture         =   "CompServicoGR.ctx":0AC5
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Numeração Automática"
      Top             =   195
      Width           =   345
   End
   Begin MSMask.MaskEdBox MaskSolicitacao 
      Height          =   315
      Left            =   1125
      TabIndex        =   0
      Top             =   195
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   300
      Left            =   3105
      TabIndex        =   1
      Top             =   195
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   90
      TabIndex        =   42
      Top             =   690
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Etapas do Serviço"
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
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   7575
      TabIndex        =   113
      Top             =   3645
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
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
      Index           =   20
      Left            =   4305
      TabIndex        =   57
      Top             =   255
      Width           =   615
   End
   Begin VB.Label LabelSolicitacao 
      AutoSize        =   -1  'True
      Caption         =   "Solicitação:"
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
      Left            =   90
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   40
      Top             =   255
      Width           =   1020
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
      Left            =   2400
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   39
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "CompServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Inicio Daniel em 08/11/2001

Option Explicit

'Declaração do GridItemServico do TAB de DadosPrincipais
Public objGridItemServico As New AdmGrid

Dim iGrid_ItemServico_Col As Integer
Dim iGrid_DescItemServico_Col As Integer
Dim iGrid_DataPrev_Col As Integer
Dim iGrid_HoraPrev_Col As Integer
Dim iGrid_DataInicio_Col As Integer
Dim iGrid_HoraInicio_Col As Integer
Dim iGrid_DataFim_Col As Integer
Dim iGrid_HoraFim_Col As Integer
'OBS -> DI = Documento Interno
Dim iGrid_DITipo_Col As Integer
Dim iGrid_DINumero_Col As Integer
Dim iGrid_DIDataEmissao_Col As Integer
'OBS -> DE = Documento Externo
Dim iGrid_DETipo_Col As Integer
Dim iGrid_DENumero_Col As Integer
Dim iGrid_DEDataEmissao_Col As Integer
Dim iGrid_DEDataRecepcao_Col As Integer
Dim iGrid_DEHoraRecepcao_Col As Integer
Dim iGrid_PlacaCaminhao_Col As Integer
Dim iGrid_PlacaCarreta_Col As Integer
Dim iGrid_Motorista_Col As Integer
Dim iGrid_Observacao_Col As Integer

'Definições dos TABs da Tela
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Complemento = 2
Private Const TAB_EtapasServico = 3

'Constantes
'ja subiram

Public iFrameAtual As Integer
Public iAlterado As Integer

Dim giServAlterado As Integer

'Início trecho Criado por Rafael menezes em 23/09/2002
'Essa variável evitará a validação de um serviço que já é válido, i.e., evitará
'que um serviço seja considerado não-associável a um comprovante por ele já estar fechado
'quando, na verdade, ele está associado ao comprovante da tela.
Dim gsServico As String
'Fim trecho Criado por Rafael menezes em 23/09/2002


'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos browser
Private WithEvents objEventoSolicitacao As AdmEvento
Attribute objEventoSolicitacao.VB_VarHelpID = -1
Private WithEvents objEventoCompServ As AdmEvento
Attribute objEventoCompServ.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoDocumentoInt As AdmEvento
Attribute objEventoDocumentoInt.VB_VarHelpID = -1
Private WithEvents objEventoDocumentoExt As AdmEvento
Attribute objEventoDocumentoExt.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Comprovante de Serviço"
    Call Form_Load

End Function

Public Function Name() As String
    
    Name = "CompServico"

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

Private Sub BotaoDocumentos_Click()

Dim lErro As Long
Dim objDoc As New ClassDocumento
Dim colSelecao As New Collection
Dim iIndice As Integer
Dim sSelecao As String

On Error GoTo Erro_BotaoDocumentos_Click

    If GridItemServico.Row = 0 Then gError 98751
    
    'Verifica qual Documento esta com o Foco foi preenchido
    If iGrid_DITipo_Col = GridItemServico.Col Then
       
        sSelecao = "TipoDoc = ?"

        colSelecao.Add 0
                
        Call Chama_Tela("DocumentoLista", colSelecao, objDoc, objEventoDocumentoInt, sSelecao)

    ElseIf iGrid_DETipo_Col = GridItemServico.Col Then
    
        sSelecao = "TipoDoc = ?"

        colSelecao.Add 1
                
        Call Chama_Tela("DocumentoLista", colSelecao, objDoc, objEventoDocumentoExt, sSelecao)
    
    End If

    Exit Sub

Erro_BotaoDocumentos_Click:

    Select Case gErr

        Case 98751
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objCompServ As New ClassCompServ
Dim dValor As Double

On Error GoTo Erro_BotaoImprimir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Codigo foi preenchido
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 99280

    'Passa a chave para objCompServ
    objCompServ.iFilialEmpresa = giFilialEmpresa
    objCompServ.lCodigo = StrParaLong(MaskCodigo.ClipText)

    'Ler o CompServ no bd
    lErro = CF("CompServGR_Le", objCompServ)
    If lErro <> SUCESSO And lErro <> 97419 Then gError 99281
    
    'Se nao achou
    If lErro <> SUCESSO Then gError 99282
        
    lErro = objRelatorio.ExecutarDireto("Comprovante de Serviço", "", 1, "", "NCODCOMPSV", objCompServ.lCodigo)
    If lErro <> SUCESSO Then gError 99283

    'Limpa a Tela
    Call Limpa_Tela_CompServ

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 99280
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPSERV_NAO_PREENCHIDO", gErr)

        Case 99281, 99283
        
        Case 99282
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub MaskCodigoContainer_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigoContainer, iAlterado)

End Sub

Private Sub MaskDemurrage_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDemurrage, iAlterado)

End Sub

Private Sub MaskTaraContainer_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskTaraContainer, iAlterado)

End Sub

Private Sub MaskValorContainer_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskValorContainer, iAlterado)

End Sub

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
'''    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******

'--------------------------------------------------//
'Rotinas dos Controles da Tela
'que nao participam de GRID
'--------------------------------------------------//

Private Sub BotaoExcluir_Click()

Dim lErro As Long, lCodigo As Long
Dim objCompServ As New ClassCompServ

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Codigo foi preenchido
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 98518

    'Passa a chave para objCompServ
    objCompServ.iFilialEmpresa = giFilialEmpresa
    objCompServ.lCodigo = StrParaLong(MaskCodigo.ClipText)

    'Ler o CompServ no bd
    lErro = CF("CompServGR_Le", objCompServ)
    If lErro <> SUCESSO And lErro <> 97419 Then gError 98519
    
    'Se nao achou
    If lErro <> SUCESSO Then gError 98520
        
    'Confirma a exclusao do CompServ
    If Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_COMPSERV", objCompServ.lCodigo, objCompServ.iFilialEmpresa) = vbYes Then

        'Exclui o comprovante
        lErro = CF("CompServGR_Exclui", objCompServ)
        If lErro <> SUCESSO Then gError 98521

        'Fecha o comando das setas, se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        'Limpa a tela
        Call Limpa_Tela_CompServ
        
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
   
    Select Case gErr

        Case 98518
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPSERV_NAO_PREENCHIDO", gErr)

        Case 98519, 98521
        
        Case 98520
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a funcao que ira efetuar a gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 98529

    'limpa a tela apos a gravacao
    Call Limpa_Tela_CompServ

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 98529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimirComprovante_Click()
'Nao esta completa, mas ja se encontra engatilhada...
'falta a funcao que ira efetuar a gravacao

Dim lErro As Long, lComando As Long
Dim sNomeTsk As String
Dim sNome As String
Dim objRelatorio As New AdmRelatorio
Dim objCompServ As New ClassCompServ
Dim objCompServItem As New ClassCompServItem
Dim sName As String
Dim lCod As Long

On Error GoTo Erro_BotaoImprimirComprovante_Click
        
    'se nao existem linhas selecionadas no grid -> erro
    If GridItemServico.Row = 0 Then gError 98560
    
    'Verifica se Codigo foi preenchido
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 99296

    If Len(Trim(GridItemServico.TextMatrix(GridItemServico.Row, iGrid_DITipo_Col))) <> 0 Then
        'pegando a descrição
        sNome = GridItemServico.TextMatrix(GridItemServico.Row, iGrid_DITipo_Col)
    Else
        gError 99284
    End If
        
    'Chama a Funcao que vai pegar o nome do TSK
    lErro = CF("Documento_Le_NomeTsk", sNomeTsk, sNome)
    If lErro <> SUCESSO Then gError 98616

    'se o nome do relatorio nao está associado ao documento ==> erro
    If Len(sNomeTsk) = 0 Then gError 105003

    'Passa a chave para objCompServ
    objCompServ.iFilialEmpresa = giFilialEmpresa
    objCompServ.lCodigo = StrParaLong(MaskCodigo.ClipText)
    
    'Le os dados da tabela Comprovante de Servico
    lErro = CF("CompServGR_Le", objCompServ)
    If lErro <> SUCESSO And lErro <> 97419 Then gError 99294
    
    'Se não encontrar --> Erro
    If lErro = 97419 Then gError 99295
    
    objCompServItem.iCodItemServico = GridItemServico.TextMatrix(GridItemServico.Row, iGrid_ItemServico_Col)
    objCompServItem.lNumIntDocOrigem = objCompServ.lNumIntDoc
    
    lErro = CF("CompServGR_Le_CompServItem_Codigo", objCompServItem)
    If lErro <> SUCESSO Then gError 99293
        
    'Usado para selecionar o name do relatório
    If UCase(sNomeTsk) = "CONTCOEN" Then
    
        'Chama a rotina que gera o sequencial
        lErro = CF("Config_ObterNumInt_Trans", "FatConfig", "NUM_PROX_RELCOLETAENTREGA", lCod)
        If lErro <> SUCESSO Then gError 99329
        
        lErro = objRelatorio.ExecutarDireto("Controle de coleta/entrega", "", 1, sNomeTsk, "NNUMINTDOCCPSV", objCompServ.lNumIntDoc, "NNUMINTDOCCPSVIT", objCompServItem.lNumIntDoc, "NCODSEQ", CStr(lCod))
        If lErro <> SUCESSO Then gError 98203
    
    ElseIf UCase(sNomeTsk) = "NFENTCOL" Then
        
        'Chama a rotina que gera o sequencial
        lErro = CF("Config_ObterNumInt_Trans", "FatConfig", "NUM_PROX_RELNFCOLETAENTREGA", lCod)
        If lErro <> SUCESSO Then gError 99330
        
        lErro = objRelatorio.ExecutarDireto("NF de entrega/coleta", "", 1, sNomeTsk, "NNUMINTDOCCPSV", objCompServ.lNumIntDoc, "NNUMINTDOCCPSVIT", objCompServItem.lNumIntDoc, "NCODSEQ", CStr(lCod))
        If lErro <> SUCESSO Then gError 99327
        
    Else
    
        lErro = objRelatorio.ExecutarDireto("Transferência de Responsabilidade", "", 1, sNomeTsk, "NNUMINTDOCCPSV", objCompServ.lNumIntDoc, "NNUMINTDOCCPSVIT", objCompServItem.lNumIntDoc)
        If lErro <> SUCESSO Then gError 99328
        
    End If
    
    Exit Sub

Erro_BotaoImprimirComprovante_Click:

    Select Case gErr
    
        Case 98560
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_NAO_SELECIONADA_GRIDITEMSERV", gErr)
    
        Case 98203, 98616, 99293, 99294, 99327, 99328, 99329, 99330
        
        Case 99284
            Call Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_PREENCHIDO1", gErr)
        
        Case 99295
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)

        Case 99296
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPSERV_NAO_PREENCHIDO", gErr)

        Case 105003
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_RELATORIO_ASSOCIADO_DOC", gErr, sNome)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long, lCod As Long
 
On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático do CompServ.
    lErro = CompServ_Codigo_Automatico(lCod)
    If lErro <> SUCESSO Then gError 97405

    'Coloca na tela o número gerado anteriormente...
    MaskCodigo.Text = lCod

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 97405

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 98574

    'Limpa a tela
    Call Limpa_Tela_CompServ

    'Fecha Comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub
    
Erro_Botaolimpar_Click:

    Select Case gErr

        Case 98574

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim colSelecao As New Collection
Dim objCompServ As New ClassCompServ

    'Verifica se o número do CompServ foi preenchido
    If Len(Trim(MaskCodigo.ClipText)) > 0 Then objCompServ.lCodigo = StrParaLong(MaskCodigo.Text)

    'Chamada do browser
    Call Chama_Tela("CompServGRLista", colSelecao, objCompServ, objEventoCompServ)

End Sub

Private Sub LabelServico_Click(Index As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sSelecao As String
Dim objSolServ As New ClassSolicitacaoServico
Dim lNumIntDocSolServ As Long

On Error GoTo Erro_LabelServico_Click

    'Verifica se o serviço foi preenchido
    If Len(Trim(MaskServico.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", MaskServico.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 97499

        objProduto.sCodigo = sProdutoFormatado
    
    End If

    'Verifica se solicitacao foi preenchida
    If Len(Trim(MaskSolicitacao.ClipText)) = 0 Then gError 98658
        
    objSolServ.lNumero = StrParaInt(MaskSolicitacao.Text)
    objSolServ.iFilialEmpresa = giFilialEmpresa
    
    'le solserv
    lErro = CF("SolicitacaoServico_Le", objSolServ)
    If lErro <> 98085 And lErro <> SUCESSO Then gError 98721
    
    If lErro <> SUCESSO Then gError 98720
    
    lNumIntDocSolServ = objSolServ.lNumIntDoc
    
    colSelecao.Add lNumIntDocSolServ
    
    Call Chama_Tela("ProdutoNaoAtendidoLista", colSelecao, objProduto, objEventoServico)
    
    Exit Sub

Erro_LabelServico_Click:

    Select Case gErr

        Case 97499, 98721

        Case 98658
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERV_NAO_PREENCHIDA", gErr, objSolServ.lNumero, objSolServ.iFilialEmpresa)
    
        Case 98720
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERV_INEXISTENTE", gErr, MaskSolicitacao.Text, giFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub LabelSolicitacao_Click()

Dim colSelecao As New Collection
Dim objSolicitacaoServico As New ClassSolicitacaoServico

    'Verifica se o número da Solicitacao foi preenchido
    If Len(Trim(MaskSolicitacao.ClipText)) > 0 Then objSolicitacaoServico.lNumero = StrParaLong(MaskSolicitacao.Text)

    Call Chama_Tela("SolicitacaoNaoAtendidoLista", colSelecao, objSolicitacaoServico, objEventoSolicitacao)
        
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
    
    'se codigo estiver preenchido
    If Len(Trim(MaskCodigo.ClipText)) <> 0 Then
        
        'critica o valor
        lErro = Long_Critica(MaskCodigo.Text)
        If lErro <> SUCESSO Then gError 97407
    
    End If
    
    Exit Sub

Erro_MaskCodigo_Validate:

    Select Case gErr
        
        Case 97407
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Sub MaskCodigoContainer_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskValorContainer_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskValorContainer_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskValorContainer_Validate

    'se Valor Container estiver preenchido
    If Len(Trim(MaskValorContainer.ClipText)) > 0 Then
        
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(MaskValorContainer.ClipText)
        If lErro <> SUCESSO Then gError 99052
        
        'coloca no formato adequado
        MaskValorContainer.Text = Format(MaskValorContainer.Text, "standard")
    
    End If
    
    Exit Sub

Erro_MaskValorContainer_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 99052
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskDataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskDataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDataEmissao, iAlterado)

End Sub

Private Sub MaskDataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskDataEmissao_Validate
    
    'se data emissao estiver preenchida
    If Len(Trim(MaskDataEmissao.ClipText)) > 0 Then
        
        'critica a data
        lErro = Data_Critica(MaskDataEmissao.Text)
        If lErro <> SUCESSO Then gError 97410
        
    End If
    
    Exit Sub

Erro_MaskDataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 97410

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskQuantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskQuantidade_Validate

    'Se quantidade estiver preenchida
    If Len(Trim(MaskQuantidade.ClipText)) > 0 Then
        
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(MaskQuantidade.ClipText)
        If lErro <> SUCESSO Then gError 97403
        
        'formata o campo
        MaskQuantidade.Text = Format(MaskQuantidade.Text, "standard")
    
    End If
    
    Exit Sub

Erro_MaskQuantidade_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 97403
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub maskQuantMaterial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskQuantMaterial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskQuantMaterial_Validate

    'Se quantidade de materiais estiver preenchida
    If Len(Trim(MaskQuantMaterial.ClipText)) > 0 Then
        
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(MaskQuantMaterial.ClipText)
        If lErro <> SUCESSO Then gError 97404
        
        'formata o campo
        MaskQuantMaterial.Text = Formata_Estoque(MaskQuantMaterial.Text)
    
    End If
    
    Exit Sub

Erro_MaskQuantMaterial_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 97404
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub


End Sub

Private Sub maskServico_Change()

    iAlterado = REGISTRO_ALTERADO
    giServAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskServico_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskServico, iAlterado)

End Sub

Private Sub maskServico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sServ As String

'retirar essa variavel
Dim objServico As New ClassServico

Dim objSolServ As New ClassSolicitacaoServico

'retirar essa variavel
Dim objTabPreco As New ClassTabPreco

Dim objCompServ As New ClassCompServ

On Error GoTo Erro_MaskServico_Validate

    If Len(Trim(MaskServico.ClipText)) = 0 Then Exit Sub
    
    'inicio trecho adicionado Rafael Menezes em 23/09/2002
    'se o serviço for igual ao global, i.e., realmente está associado à solicitação da tela, sai
    If MaskServico.ClipText = gsServico Then Exit Sub
    'fim trecho adicionado Rafael Menezes em 23/09/2002
    
    'Como nao existe servico para uma solicitacao inexistente...
    If Len(Trim(MaskSolicitacao.ClipText)) = 0 Then gError 98616
    
    objSolServ.lNumero = MaskSolicitacao.Text
    objSolServ.iFilialEmpresa = giFilialEmpresa
    
    'início trecho adicionado por rafael menezes em 23/09/2002
    lErro = CF("Produto_Formata", MaskServico.Text, sServ, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 109110
    
    objCompServ.sProduto = sServ
    
    'verifica se o código digitado como serviço já não está associado a um comprovante dakela solicitação
    lErro = CF("ServicosAssociadosComprovantes_Le", objSolServ, objCompServ)
    If lErro <> SUCESSO And lErro <> 109103 And lErro <> 109106 Then gError 109107
    
    'se não encontrou a solicitação-> erro
    If lErro = 109103 Then gError 109108
    
    'se encontrou a associação entre a solicitação e o comprovante-> erro
    If lErro = SUCESSO Then gError 109109
    
    'essa parte do código continuará se o erro gerado na chamada de ServicosAssociadosComprovantes_Le for 109106
    'fim trecho adicionado por rafael menezes em 23/09/2002
    
    'garanto q a solicitacao existe, pois ja foi passado o validate dela
    lErro = CF("SolicitacaoServicoNaoAtendidos_Le", objSolServ)
    If lErro <> SUCESSO And lErro <> 99322 Then gError 98740

    lErro = CF("Produto_Formata", MaskServico.Text, sServ, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 98741
    
    objProduto.sCodigo = sServ
    
    lErro = CF("Verifica_Serv_SolServ", objSolServ, objProduto)
    If lErro <> SUCESSO Then gError 98735
    
    'se o servico foi alterado
    If giServAlterado = REGISTRO_ALTERADO Then
    
       'Se o Serviço está Preenchido...
       If Len(Trim(MaskServico.ClipText)) <> 0 Then
        
           'Traz produto pra tela
           lErro = Traz_Produto_Tela(objProduto)
           If lErro <> SUCESSO Then gError 98608
           
                  
       Else
        
          'Limpa a descricao
          labelDescricao.Caption = ""
                
          'Limpa o Grid
          Call Grid_Limpa(objGridItemServico)
        
       End If

    End If
    
    gsServico = MaskServico.ClipText
    giServAlterado = 0
    
    Exit Sub

Erro_MaskServico_Validate:

    Cancel = True

    Select Case gErr

        Case 98608, 98735, 98740, 98741, 109107, 109110
            giServAlterado = 0
            
        Case 98616
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_SERVICO_PARA_SOLICITACAO_VAZIA", gErr)
            giServAlterado = REGISTRO_ALTERADO
            
        Case 109108
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAOSERVICO_NAO_CADASTRADA", gErr, objSolServ.iFilialEmpresa, objSolServ.lNumero)
        
        Case 109109
        'é o cara
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_JA_ASSOCIADO_COMPROVANTE", gErr, objCompServ.sProduto, objCompServ.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            giServAlterado = 0
            
    End Select
   
    Exit Sub

End Sub

Private Sub MaskSolicitacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskSolicitacao_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskSolicitacao, iAlterado)

End Sub

Private Sub MaskSolicitacao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objSolServ As New ClassSolicitacaoServico

On Error GoTo Erro_MaskSolicitacao_Validate

    'se solicitacao estiver preenchida
    If Len(Trim(MaskSolicitacao.ClipText)) <> 0 Then
        
        'critica o valor
        lErro = Long_Critica(MaskSolicitacao.Text)
        If lErro <> SUCESSO Then gError 97408
    
        'Move a chave para objSolServ
        objSolServ.lNumero = StrParaInt(MaskSolicitacao.Text)
        objSolServ.iFilialEmpresa = giFilialEmpresa
        
        'traz a solicitacao pra tela
        lErro = Traz_SolicitacaoServico_Tela(objSolServ)
        If lErro <> SUCESSO Then gError 98603
       
    End If
    
    Exit Sub

Erro_MaskSolicitacao_Validate:

    Cancel = True

    Select Case gErr
        
        Case 97408, 98603
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Sub MaskTaraContainer_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskTaraContainer_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskTaraContainer_Validate

    'Se a Tara estiver preenchida
    If Len(Trim(MaskTaraContainer.ClipText)) > 0 Then
        
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(MaskTaraContainer.ClipText)
        If lErro <> SUCESSO Then gError 97415
        
        'Formata o campo
        MaskTaraContainer.Text = Formata_Estoque(MaskTaraContainer.Text)
    
    End If
    
    Exit Sub

Erro_MaskTaraContainer_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 97415
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
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
        If lErro <> SUCESSO Then gError 97405
        
        'Converte o valor para double
        dValorMaterial = StrParaDbl(MaskValorMerc.Text)
        
        'formata o valor e coloca o mesmo na tela
        MaskValorMerc.Text = Format(dValorMaterial, "Standard")
    
    End If
    
    Exit Sub

Erro_MaskValorMerc_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 97405
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se o Frame atual não corresponde ao TAB clicado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame selecionado visível
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
    
    End If

End Sub

Private Sub TextLacreContainer_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextObs_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextUM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_DownClick

    'decrementa a data
    lErro = Data_Up_Down_Click(MaskDataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 97411

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case gErr

        Case 97411
            MaskDataEmissao.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_UpClick

    'incrementa a data
    lErro = Data_Up_Down_Click(MaskDataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 97412

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case gErr

        Case 97412
            MaskDataEmissao.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'--------------------------------------------------//
'Rotinas dos Controles da Tela
'que participam de GRID
'--------------------------------------------------//

Public Sub GridItemServico_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItemServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItemServico, iAlterado)
    End If

End Sub

Public Sub GridItemServico_GotFocus()
    Call Grid_Recebe_Foco(objGridItemServico)
End Sub

Public Sub GridItemServico_EnterCell()
    Call Grid_Entrada_Celula(objGridItemServico, iAlterado)
End Sub

Public Sub GridItemServico_LeaveCell()
    Call Saida_Celula(objGridItemServico)
End Sub

Public Sub GridItemServico_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItemServico)
End Sub

Public Sub GridItemServico_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItemServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItemServico, iAlterado)
    End If

End Sub

Public Sub GridItemServico_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItemServico)
End Sub

Public Sub GridItemServico_RowColChange()
    
    Call Grid_RowColChange(objGridItemServico)
    
End Sub

Public Sub GridItemServico_Scroll()
    Call Grid_Scroll(objGridItemServico)
End Sub

Private Sub MaskItemServico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub MaskItemServico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub MaskItemServico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskItemServico
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub textDescItemServico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub textDescItemServico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub textDescItemServico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = TextDescItemServico
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDataPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDataPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDataPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDataPrev
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskHoraPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskHoraPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskHoraPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskHoraPrev
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDataInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDataInicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDataInicio
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskHoraInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskHoraInicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskHoraInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskHoraInicio
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDataFim_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDataFim_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDataFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDataFim
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskHoraFim_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskHoraFim_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskHoraFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskHoraFim
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDocIntTipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDocIntTipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDocIntTipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDocIntTipo
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDocIntNumero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDocIntNumero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDocIntNumero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDocIntNumero
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDocIntDataEmi_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDocIntDataEmi_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDocIntDataEmi_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDocIntDataEmi
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDocExtTipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDocExtTipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDocExtTipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDocExtTipo
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDocExtNumero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDocExtNumero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDocExtNumero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDocExtNumero
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDocExtDataEmi_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDocExtDataEmi_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDocExtDataEmi_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDocExtDataEmi
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDocExtDataRec_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDocExtDataRec_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDocExtDataRec_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDocExtDataRec
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskDocExtHoraRec_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskDocExtHoraRec_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskDocExtHoraRec_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskDocExtHoraRec
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaskPlacaCaminhao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub MaskPlacaCaminhao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub MaskPlacaCaminhao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskPlacaCaminhao
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaskPlacaCarreta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub MaskPlacaCarreta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub MaskPlacaCarreta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskPlacaCarreta
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub maskMotorista_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub maskMotorista_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub maskMotorista_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = MaskMotorista
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub textObservacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItemServico)

End Sub

Private Sub textObservacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItemServico)

End Sub

Private Sub textObservacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItemServico.objControle = TextObservacao
    lErro = Grid_Campo_Libera_Foco(objGridItemServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'--------------------------------------------------//
'Rotinas referentes as saidas de celula
'--------------------------------------------------//

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        'Verifica qual celula do gr0id esta deixando
        'de ser a corrente para chamar a funcao de
        'saida celula adequada...
        Select Case objGridInt.objGrid.Col

            Case iGrid_DataFim_Col
                
                lErro = Saida_Celula_DataFim(objGridInt)
                If lErro <> SUCESSO Then gError 97429
                
            Case iGrid_DataInicio_Col
                
                lErro = Saida_Celula_DataInicio(objGridInt)
                If lErro <> SUCESSO Then gError 97430
                
            Case iGrid_DataPrev_Col
                
                lErro = Saida_Celula_DataPrev(objGridInt)
                If lErro <> SUCESSO Then gError 97431
                
            Case iGrid_DEDataEmissao_Col
                
                lErro = Saida_Celula_DEDataEmissao(objGridInt)
                If lErro <> SUCESSO Then gError 97432
        
            Case iGrid_DEDataRecepcao_Col
        
                lErro = Saida_Celula_DEDataRecepcao(objGridInt)
                If lErro <> SUCESSO Then gError 97433
                
            Case iGrid_DEHoraRecepcao_Col
                
                lErro = Saida_Celula_DEHoraRecepcao(objGridInt)
                If lErro <> SUCESSO Then gError 97434
                
            Case iGrid_DENumero_Col
            
                lErro = Saida_Celula_DENumero(objGridInt)
                If lErro <> SUCESSO Then gError 97435
                
            Case iGrid_DETipo_Col
            
                lErro = Saida_Celula_DETipo(objGridInt)
                If lErro <> SUCESSO Then gError 97436
                
            Case iGrid_DIDataEmissao_Col
            
                lErro = Saida_Celula_DIDataEmissao(objGridInt)
                If lErro <> SUCESSO Then gError 97437
                
            Case iGrid_DINumero_Col
            
                lErro = Saida_Celula_DINumero(objGridInt)
                If lErro <> SUCESSO Then gError 97438
                
            Case iGrid_DITipo_Col
            
                lErro = Saida_Celula_DITipo(objGridInt)
                If lErro <> SUCESSO Then gError 97439
                
            Case iGrid_HoraFim_Col
            
                lErro = Saida_Celula_HoraFim(objGridInt)
                If lErro <> SUCESSO Then gError 97440
                
            Case iGrid_HoraInicio_Col
            
                lErro = Saida_Celula_HoraInicio(objGridInt)
                If lErro <> SUCESSO Then gError 97441
                
            Case iGrid_HoraPrev_Col
            
                lErro = Saida_Celula_HoraPrev(objGridInt)
                If lErro <> SUCESSO Then gError 97442
                
            Case iGrid_Motorista_Col
            
                lErro = Saida_Celula_Motorista(objGridInt)
                If lErro <> SUCESSO Then gError 97443
                
            Case iGrid_Observacao_Col
            
                lErro = Saida_Celula_Observacao(objGridInt)
                If lErro <> SUCESSO Then gError 97444
                
            Case iGrid_PlacaCaminhao_Col
            
                lErro = Saida_Celula_PlacaCaminhao(objGridInt)
                If lErro <> SUCESSO Then gError 97445
                
            Case iGrid_PlacaCarreta_Col
            
                lErro = Saida_Celula_PlacaCarreta(objGridInt)
                If lErro <> SUCESSO Then gError 97446
               
        End Select

    End If

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97447
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 97429 To 97446
        
        Case 97447
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataFim(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataFim que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataFim

    Set objGridInt.objControle = MaskDataFim

    If Len(Trim(MaskDataFim.ClipText)) > 0 Then
        lErro = Data_Critica(MaskDataFim.Text)
        If lErro <> SUCESSO Then gError 98660
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97448

    Saida_Celula_DataFim = SUCESSO

    Exit Function

Erro_Saida_Celula_DataFim:

    Saida_Celula_DataFim = gErr

    Select Case gErr

        Case 97448, 98660
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataInicio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataInicio que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataInicio

    Set objGridInt.objControle = MaskDataInicio

    If Len(Trim(MaskDataInicio.ClipText)) > 0 Then
        lErro = Data_Critica(MaskDataInicio.Text)
        If lErro <> SUCESSO Then gError 98661
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97449

    Saida_Celula_DataInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_DataInicio:

    Saida_Celula_DataInicio = gErr

    Select Case gErr

        Case 97449, 98661
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrev(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataPrev que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrev

    Set objGridInt.objControle = MaskDataPrev

    If Len(Trim(MaskDataPrev.ClipText)) > 0 Then
        lErro = Data_Critica(MaskDataPrev.Text)
        If lErro <> SUCESSO Then gError 98662
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97450

    Saida_Celula_DataPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrev:

    Saida_Celula_DataPrev = gErr

    Select Case gErr

        Case 97450, 98662
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DEDataEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DEDataEmissao que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DEDataEmissao

    Set objGridInt.objControle = MaskDocExtDataEmi

    If Len(Trim(MaskDocExtDataEmi.ClipText)) > 0 Then
        lErro = Data_Critica(MaskDocExtDataEmi.Text)
        If lErro <> SUCESSO Then gError 98663
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97451

    Saida_Celula_DEDataEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_DEDataEmissao:

    Saida_Celula_DEDataEmissao = gErr

    Select Case gErr

        Case 97451, 98663
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DEDataRecepcao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DEDataRecepcao que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DEDataRecepcao

    Set objGridInt.objControle = MaskDocExtDataRec

    If Len(Trim(MaskDocExtDataRec.ClipText)) > 0 Then
        lErro = Data_Critica(MaskDocExtDataRec.Text)
        If lErro <> SUCESSO Then gError 98664
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97452

    Saida_Celula_DEDataRecepcao = SUCESSO

    Exit Function

Erro_Saida_Celula_DEDataRecepcao:

    Saida_Celula_DEDataRecepcao = gErr

    Select Case gErr

        Case 97452, 98664
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DEHoraRecepcao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DEHoraRecepcao que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DEHoraRecepcao

    Set objGridInt.objControle = MaskDocExtHoraRec

    If Len(Trim(MaskDocExtHoraRec.ClipText)) > 0 Then
        lErro = Hora_Critica(MaskDocExtHoraRec.Text)
        If lErro <> SUCESSO Then gError 98665
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97453

    Saida_Celula_DEHoraRecepcao = SUCESSO

    Exit Function

Erro_Saida_Celula_DEHoraRecepcao:

    Saida_Celula_DEHoraRecepcao = gErr

    Select Case gErr

        Case 97453, 98665
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DENumero(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DENumero que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DENumero

    Set objGridInt.objControle = MaskDocExtNumero

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97454

    Saida_Celula_DENumero = SUCESSO

    Exit Function

Erro_Saida_Celula_DENumero:

    Saida_Celula_DENumero = gErr

    Select Case gErr

        Case 97454
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DETipo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DETipo que está deixando de ser a corrente

Dim lErro As Long
Dim objDoc As New ClassDocumento

On Error GoTo Erro_Saida_Celula_DETipo

    Set objGridInt.objControle = MaskDocExtTipo

    '??? arrumar os tratamentos de erro --> ok!!!
    lErro = CF("TP_Doc_Le", objDoc, MaskDocExtTipo.Text)
    If lErro <> SUCESSO And lErro <> 98550 And lErro <> 98552 And lErro <> 98180 Then gError 98557

    If lErro = 98550 Then gError 98684
    
    If lErro = 98552 Then gError 98685
    
    If lErro <> 98180 Then
    
        'testa se o tipo do documento eh externo..
        If objDoc.iTipoDoc = DOCUMENTO_INTERNO Then gError 98950
        
        MaskDocExtTipo.Text = objDoc.sNomeReduzido
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97455

    Saida_Celula_DETipo = SUCESSO

    Exit Function

Erro_Saida_Celula_DETipo:

    Saida_Celula_DETipo = gErr

    Select Case gErr

        Case 97455, 98557
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 98684
            'Envia aviso que Documento não existe e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_DOCUMENTO2", objDoc.sNomeReduzido)
    
            If lErro = vbYes Then
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                'Chama tela de Documento
                lErro = Chama_Tela("Documento", objDoc)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItemServico)
            
            End If

        Case 98685
            'Envia aviso que Documento não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_DOCUMENTO", objDoc.iCodigo)
    
                If lErro = vbYes Then
                    
                    Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                    
                    'Chama tela de Documento
                    lErro = Chama_Tela("Documento", objDoc)
                Else
                    Call Grid_Trata_Erro_Saida_Celula(objGridItemServico)
                
                End If
        
        Case 98950
            MaskDocExtTipo.Text = ""
            
            Call Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_EH_EXTERNO", gErr, objDoc.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DIDataEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DIDataEmissao que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DIDataEmissao

    Set objGridInt.objControle = MaskDocIntDataEmi

    If Len(Trim(MaskDocIntDataEmi.ClipText)) > 0 Then
        
        lErro = Data_Critica(MaskDocIntDataEmi.Text)
        If lErro <> SUCESSO Then gError 98666
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97456

    Saida_Celula_DIDataEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_DIDataEmissao:

    Saida_Celula_DIDataEmissao = gErr

    Select Case gErr

        Case 97456, 98666
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DINumero(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DINumero que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DINumero

    Set objGridInt.objControle = MaskDocIntNumero

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97457

    Saida_Celula_DINumero = SUCESSO

    Exit Function

Erro_Saida_Celula_DINumero:

    Saida_Celula_DINumero = gErr

    Select Case gErr

        Case 97457
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DITipo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DINumero que está deixando de ser a corrente

Dim lErro As Long
Dim objDoc As New ClassDocumento

On Error GoTo Erro_Saida_Celula_DITipo

    Set objGridInt.objControle = MaskDocIntTipo

    lErro = CF("TP_Doc_Le", objDoc, MaskDocIntTipo.Text)
    If lErro <> SUCESSO And lErro <> 98550 And lErro <> 98552 And lErro <> 98180 Then gError 98556

    If lErro = 98550 Then gError 98686
    
    If lErro = 98552 Then gError 98687

    If lErro <> 98180 Then

        'testa se o tipo do documento eh externo..
        If objDoc.iTipoDoc = DOCUMENTO_EXTERNO Then gError 98951
    
        MaskDocIntTipo.Text = objDoc.sNomeReduzido

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97492

    Saida_Celula_DITipo = SUCESSO

    Exit Function

Erro_Saida_Celula_DITipo:

    Saida_Celula_DITipo = gErr

    Select Case gErr

        Case 97492, 98556
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 98686
            'Envia aviso que Documento não existe e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_DOCUMENTO2", objDoc.sNomeReduzido)
    
            If lErro = vbYes Then
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                'Chama tela de Documento
                lErro = Chama_Tela("Documento", objDoc)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItemServico)
            
            End If

        Case 98687
            'Envia aviso que Documento não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_DOCUMENTO", objDoc.iCodigo)
    
                If lErro = vbYes Then
                    
                    Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                    
                    'Chama tela de Documento
                    lErro = Chama_Tela("Documento", objDoc)
                Else
                    Call Grid_Trata_Erro_Saida_Celula(objGridItemServico)
                
                End If

        Case 98951
            MaskDocIntTipo.Text = ""
            
            Call Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_EH_INTERNO", gErr, objDoc.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_HoraFim(objGridInt As AdmGrid) As Long
'Faz a crítica da célula HoraFim que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_HoraFim

    Set objGridInt.objControle = MaskHoraFim

    If Len(Trim(MaskHoraFim.ClipText)) > 0 Then
        lErro = Hora_Critica(MaskHoraFim.Text)
        If lErro <> SUCESSO Then gError 98667
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97458

    Saida_Celula_HoraFim = SUCESSO

    Exit Function

Erro_Saida_Celula_HoraFim:

    Saida_Celula_HoraFim = gErr

    Select Case gErr

        Case 97458, 98667
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_HoraInicio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula HoraInicio que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_HoraInicio

    Set objGridInt.objControle = MaskHoraInicio

    If Len(Trim(MaskHoraInicio.ClipText)) > 0 Then
        lErro = Hora_Critica(MaskHoraInicio.Text)
        If lErro <> SUCESSO Then gError 98668
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97459

    Saida_Celula_HoraInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_HoraInicio:

    Saida_Celula_HoraInicio = gErr

    Select Case gErr

        Case 97459, 98668
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_HoraPrev(objGridInt As AdmGrid) As Long
'Faz a crítica da célula HoraPrev que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_HoraPrev

    Set objGridInt.objControle = MaskHoraPrev

    If Len(Trim(MaskHoraPrev.ClipText)) > 0 Then
        lErro = Hora_Critica(MaskHoraPrev.Text)
        If lErro <> SUCESSO Then gError 98670
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97460

    Saida_Celula_HoraPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_HoraPrev:

    Saida_Celula_HoraPrev = gErr

    Select Case gErr

        Case 97460, 98670
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Motorista(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Motorista que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Motorista

    Set objGridInt.objControle = MaskMotorista

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97461

    Saida_Celula_Motorista = SUCESSO

    Exit Function

Erro_Saida_Celula_Motorista:

    Saida_Celula_Motorista = gErr

    Select Case gErr

        Case 97461
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Observacao que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = TextObservacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97462

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 97462
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PlacaCaminhao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula PlacaCaminhao que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PlacaCaminhao

    Set objGridInt.objControle = MaskPlacaCaminhao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97463

    Saida_Celula_PlacaCaminhao = SUCESSO

    Exit Function

Erro_Saida_Celula_PlacaCaminhao:

    Saida_Celula_PlacaCaminhao = gErr

    Select Case gErr

        Case 97463
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PlacaCarreta(objGridInt As AdmGrid) As Long
'Faz a crítica da célula PlacaCarreta que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PlacaCarreta

    Set objGridInt.objControle = MaskPlacaCarreta

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 97464

    Saida_Celula_PlacaCarreta = SUCESSO

    Exit Function

Erro_Saida_Celula_PlacaCarreta:

    Saida_Celula_PlacaCarreta = gErr

    Select Case gErr

        Case 97464
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

'--------------------------------------------------//
'Eventos Selecao dos Browsers
'--------------------------------------------------//

Private Sub objEventoDocumentoInt_evSelecao(obj1 As Object)

Dim objDoc As ClassDocumento
Dim lErro As Long

On Error GoTo Erro_objEventoDocumento_evSelecao

    Set objDoc = obj1

    'Traz o documento interno selecionado pra tela
    GridItemServico.TextMatrix(GridItemServico.Row, iGrid_DITipo_Col) = objDoc.sNomeReduzido
    MaskDocIntTipo.Text = objDoc.sNomeReduzido
        
        
    MaskDocIntTipo = objDoc.sNomeReduzido
    
    Me.Show

    Exit Sub

Erro_objEventoDocumento_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoDocumentoExt_evSelecao(obj1 As Object)

Dim objDoc As ClassDocumento
Dim lErro As Long

On Error GoTo Erro_objEventoDocumento_evSelecao

    Set objDoc = obj1

    'Traz o documento externo selecionado pra tela
    GridItemServico.TextMatrix(GridItemServico.Row, iGrid_DETipo_Col) = objDoc.sNomeReduzido
    MaskDocExtTipo = objDoc.sNomeReduzido
        
    Me.Show

    Exit Sub

Erro_objEventoDocumento_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCompServ_evSelecao(obj1 As Object)

Dim objCompServ As ClassCompServ
Dim lErro As Long

On Error GoTo Erro_objEventoCompServ_evSelecao

    Set objCompServ = obj1

    'Traz o compserv selecionado pra tela
    lErro = Traz_CompServ_Tela(objCompServ)
    If lErro <> SUCESSO And lErro <> 97466 Then gError 97494
   
    If lErro <> SUCESSO Then gError 97493
   
    Me.Show

    Exit Sub

Erro_objEventoCompServ_evSelecao:

    Select Case gErr

        Case 97493
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)

        Case 97494
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim lErro As Long
Dim sServico As String

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'preenche os dados referentes ao produto
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sServico)
    If lErro <> SUCESSO Then gError 98675
   
    'Início trecho adicionado por RAfael Menezes em 23/09/2002
    gsServico = sServico
    'Fim trecho adicionado por RAfael Menezes em 23/09/2002
    
    MaskServico.PromptInclude = False
    MaskServico.Text = sServico
    MaskServico.PromptInclude = True
    
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO And lErro <> 97480 Then gError 97496

    If lErro <> SUCESSO Then gError 97495

    'Fecha o comando das setas, se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 97495
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 97496
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSolicitacao_evSelecao(obj1 As Object)

Dim objSolServ As ClassSolicitacaoServico
Dim lErro As Long

On Error GoTo Erro_objEventoSolicitacao_evSelecao

    Set objSolServ = obj1

    'Traz a solicitacao pra tela...
    lErro = Traz_SolicitacaoServico_Tela(objSolServ)
    If lErro <> SUCESSO Then gError 97498
   
    'Fecha o comando das setas, se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoSolicitacao_evSelecao:

    Select Case gErr

        Case 97497
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERV_NAO_CADASTRADA", gErr, objSolServ.lNumero, objSolServ.iFilialEmpresa)

        Case 97498
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'--------------------------------------------------//
'Funcoes de Leitura
'--------------------------------------------------//

'ja subiram....

'--------------------------------------------------//
'Funcoes de Gravacao
'--------------------------------------------------//

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCompServ As New ClassCompServ
Dim iIndex As Integer

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se a Solicitacao esta preenchida
    If Len(Trim(MaskSolicitacao.ClipText)) = 0 Then gError 98534
    
    'Verifica se Codigo da Comprovante está preenchido
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 98531

    'Verifica se tipo de status esta preenchido
    If ComboStatus.ListIndex < 0 Then gError 98535
    
    'Verifica se o Servico esta preenchido
    If Len(Trim(MaskServico.ClipText)) = 0 Then gError 98532

    'Verifica se a quantidade do servico esta preenchida
    If Len(Trim(MaskQuantidade.ClipText)) = 0 Then gError 98533
    
    If ComboStatus.ListIndex = STATUS_FATURAVEL Or ComboStatus.ListIndex = STATUS_CONCLUIDO Then
        
        'Verifica se o material esta preenchido
        If Len(Trim(LabelMaterial.Caption)) = 0 Then gError 99311
    
        'Verifica se a quantidade de material esta preenchida
        If Len(Trim(MaskQuantMaterial.ClipText)) = 0 Or StrParaDbl(MaskQuantMaterial.Text) = 0 Then gError 98536
    
        'Verifica se a UM esta preenchida
        If Len(Trim(TextUM.Text)) = 0 Then gError 98537
        
        'Verifica se a quantidade da mercadoria esta preenchida
        If Len(Trim(MaskValorMerc.ClipText)) = 0 Or StrParaDbl(MaskValorMerc.Text) = 0 Then gError 99312
    
        'Verifica se o codigo do container esta preenchido
        If Len(Trim(MaskCodigoContainer.ClipText)) = 0 Then gError 98538
    
        'Verifica se o Valor do container esta preenchido
        If Len(Trim(MaskValorContainer.ClipText)) = 0 Or StrParaDbl(MaskValorContainer.Text) = 0 Then gError 99053
    
        'Verifica se a Tara do container esta preenchida
        If Len(Trim(MaskTaraContainer.ClipText)) = 0 Or StrParaDbl(MaskTaraContainer.Text) = 0 Then gError 98539
    
        'Verifica se o Lacre do Container esta preenchido
        If Len(Trim(TextLacreContainer.Text)) = 0 Then gError 98540
     End If

'    'Verifica se o grid possui pelo menos 1 item
    If objGridItemServico.iLinhasExistentes = 0 Then gError 98705
           
    'Se o navio nao estiver preenchido, critica a data do demurrage
    If Len(Trim(LabelProgNavio.Caption)) <= 0 Then
        If Len(Trim(MaskDemurrage.ClipText)) > 0 Then gError 98750
    End If
    
    'Verifica consistencia da tela (horas e datas)
    lErro = Testa_Consistencia_Tela
    If lErro <> SUCESSO Then gError 98713
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objCompServ)
    If lErro <> SUCESSO Then gError 98541

    lErro = Trata_Alteracao(objCompServ, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98752
    
    'Chama a funcao que vai efetuar, efetivamente, a gravacao
    lErro = CF("CompServGR_Grava", objCompServ)
    If lErro <> SUCESSO Then gError 98542
    
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

        Case 98531
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPSERV_NAO_PREENCHIDO", gErr)

        Case 98532
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO", gErr)

        Case 98533
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDSERVICO_NAO_PREENCHIDA", gErr)

        Case 98534
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERV_NAO_PREENCHIDA", gErr)

        Case 98535
            Call Rotina_Erro(vbOKOnly, "ERRO_STATUS_NAO_PREENCHIDO", gErr)
            
        Case 98536
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDMATERIAL_NAO_PREENCHIDA", gErr)
            
        Case 98537
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDO", gErr)
        
        Case 98538
            Call Rotina_Erro(vbOKOnly, "ERRO_CODCONTAINER_NAO_PREENCHIDO", gErr)
        
        Case 98539
            Call Rotina_Erro(vbOKOnly, "ERRO_TARA_NAO_PREENCHIDA", gErr)
        
        Case 98540
            Call Rotina_Erro(vbOKOnly, "ERRO_LACRE_NAO_PREENCHIDO", gErr)
        
        Case 98541, 98542, 98713, 98752
                
        Case 98705
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_POSSUI_ITENS", gErr)
                
        Case 98750
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADEMURRAGE_SEM_PROGNAVIO", gErr)
        
        Case 99053
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORCONTAINER_NAO_PREENCHIDO", gErr)
            
        Case 99311
            Call Rotina_Erro(vbOKOnly, "ERRO_MATERIAL_NAO_PREENCHIDO", gErr)
        
        Case 99312
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORMERC_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function
    
End Function

'--------------------------------------------------//
'Funcoes de Exclusao
'--------------------------------------------------//

'ja subiram...

'--------------------------------------------------//
'Funcoes que fazem lock
'--------------------------------------------------//

'ja subiram

'--------------------------------------------------//
'Funcoes que trazem dados para a tela
'--------------------------------------------------//

Private Function Move_Dados_Tela(objCompServ As ClassCompServ) As Long
'Move os dados carregados em objCompServ para a tela

Dim lErro As Long
Dim objSolServ As New ClassSolicitacaoServico
Dim objProduto As New ClassProduto
Dim sServico As String

On Error GoTo Erro_Move_Dados_Tela
        
    'colocando o cod do compserv
    MaskCodigo.Text = objCompServ.lCodigo
       
    'passando a chave...
    objSolServ.lNumIntDoc = objCompServ.lNumIntDocOrigem
    
    'pegando o código da solicitacao para a filial em questao
    lErro = CF("Solicitacao_Servico_Le_NumIntDoc", objSolServ)
    If lErro <> SUCESSO And lErro <> 97472 Then gError 98473
    
    If lErro <> SUCESSO Then gError 98474
    
    'preenche os dados referentes a solicitacao
    lErro = Traz_SolicitacaoServico_Tela(objSolServ)
    If lErro <> SUCESSO Then gError 98651
    
    'colocando o codigo do produto no obj
    objProduto.sCodigo = objCompServ.sProduto
    
    'preenche os dados referentes ao produto
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sServico)
    If lErro <> SUCESSO Then gError 98675
    
    'Início trecho adicionado por RAfael Menezes em 23/09/2002
    gsServico = sServico
    'Início trecho adicionado por RAfael Menezes em 23/09/2002
    
    MaskServico.PromptInclude = False
    MaskServico.Text = sServico
    MaskServico.PromptInclude = True
    
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then gError 98652
    
    '**********************************************************
    'observacao... alguns campos serao sobrescritos em virtude
    'de serem trazidos pelas funcoes traz_xxxx_tela.
    'As funcoes traz_xxxx_tela foram aproveitadas, pois ja trazem
    'varios dados relevantes para a tela.
    '**********************************************************
    
    'preenchendo a combo de status
    ComboStatus.ListIndex = objCompServ.iSituacao
    
    'Move a data de emissao do compserv
    MaskDataEmissao.PromptInclude = False
    MaskDataEmissao.Text = Format(objCompServ.dtDataEmissao, "dd/mm/yy")
    MaskDataEmissao.PromptInclude = True
    
    'Move a qtd material, quantidade, UM, valor mercadoria e pedagio
    '?????? Formatar. Ver outras telas.
    MaskQuantMaterial.Text = Format(objCompServ.dQuantMaterial, "standard")
    '?????? Formatar. Ver outras telas.
    MaskQuantidade.Text = Format(objCompServ.dQuantidade, "standard")
    TextUM.Text = objCompServ.sUM
    MaskValorMerc.Text = Format(objCompServ.dValorMercadoria, "standard")
    LabelPedagio.Caption = Format(objCompServ.dPedagio, "standard")
    LabelPrecoUnitario.Caption = Format(objCompServ.dFretePeso, "standard")
    '???? Formatar
    LabelAdValoren.Caption = (objCompServ.dAdValoren * 100) & "%"
    
    '????? Ver como outras telas tratam de data nula
    'move o demurrage..
    If objCompServ.dtDataDemurrage <> DATA_NULA Then
        MaskDemurrage.PromptInclude = False
        MaskDemurrage.Text = Format(objCompServ.dtDataDemurrage, "dd/mm/yy")
        MaskDemurrage.PromptInclude = True
    Else
        MaskDemurrage.PromptInclude = False
        MaskDemurrage.Text = ""
        MaskDemurrage.PromptInclude = True
    
    End If
    
    'move o numero, o valor, a tara, o lacre e a obs.
    MaskCodigoContainer.Text = objCompServ.sCodigoContainer
    MaskValorContainer.Text = Format(objCompServ.dValorContainer, "standard")
    MaskTaraContainer.Text = Format(objCompServ.dTara, "standard")
    TextLacreContainer.Text = objCompServ.sLacre
    TextObs.Text = objCompServ.sObservacao
    
    'Traz as informacoes do Grid pra tela
    lErro = Carrega_GridItemServico2(objCompServ)
    If lErro <> SUCESSO Then gError 98428
    
    Move_Dados_Tela = SUCESSO
    
    Exit Function

Erro_Move_Dados_Tela:

    Move_Dados_Tela = gErr

    Select Case gErr

        Case 98428, 98651, 98652, 98473
        
        Case 98474
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_COMPSERV_SOLSERV", gErr, objSolServ.lNumero, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
   
End Function

Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long
'Traz o Produto do obj para tela fazendo a formatacao do codigo

Dim lErro As Long
Dim iServicoPreenchido As Integer
Dim sServico As String
Dim objSolServServico As New ClassServico
Dim objTabPrecoItem As New ClassTabPrecoItens
Dim objSolServ As New ClassSolicitacaoServico
Dim dPreco As Double, dQtd As Double

On Error GoTo Erro_Traz_Produto_Tela

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", MaskServico.Text, objProduto, iServicoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 97479
    
    If lErro = 51381 Then gError 97480

    'carrega os dados necessarios em objsolserv..
    objSolServ.lNumero = MaskSolicitacao.ClipText
    objSolServ.iFilialEmpresa = giFilialEmpresa
    
    'Le a quantidade e preco do produto
    lErro = CF("SolServServTabPreItens_Le", objSolServ, objProduto, dPreco, dQtd)
    If lErro <> SUCESSO Then gError 98621
        
    'Verifica se é de Faturamento
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 97484
    
    'coloca os valores lidos na tela...
    labelDescricao.Caption = objProduto.sDescricao
    
    If Len(Trim(MaskQuantidade.ClipText)) = 0 Then MaskQuantidade.Text = QUANTIDADE_DEFAULT
    If Len(Trim(LabelPrecoUnitario.Caption)) = 0 Then LabelPrecoUnitario.Caption = Format(dPreco, "standard")
        
    'move a chave primario do produto para um objeto da classe
    'servico
    objSolServServico.sProduto = objProduto.sCodigo
    
    'carrega o grid com o objeto da classe servico
    Call Carrega_GridItemServico(objSolServServico)
        
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 97479, 98621

        Case 97480
            'nao achou o produto...
            'Possivel inconsistencia no relacionamento da tabela
            'Solicitacao de Servico e Servico
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_RELACIONADO_COM_SOLSERV", gErr, objProduto.sCodigo, MaskSolicitacao.Text, giFilialEmpresa)
            
            'Início trecho adicionado por RAfael Menezes em 23/09/2002
            gsServico = ""
            'Fim trecho adicionado por RAfael Menezes em 23/09/2002
            
            MaskServico.PromptInclude = False
            MaskServico.Text = ""
            MaskServico.PromptInclude = True
            labelDescricao.Caption = ""
            
            MaskQuantidade.Text = ""
            LabelPrecoUnitario.Caption = ""

       Case 97482
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PODE_SER_VENDIDO", gErr, objProduto.sCodigo)

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Function Traz_OriDest_Tela(objSolServ As ClassSolicitacaoServico) As Long
'Traz a Origem e o Destino para tela a partir de um objSolServ
'Traz tambem os dados referentes a tabela de preco, pois esses sao
'necessarios na tela e sera preciso passar pela tabela de preco
'para poder chegar na Origem e no Destino

Dim lErro As Long
Dim objOriDest As New ClassOrigemDestino
Dim objTabPreco As New ClassTabPreco

On Error GoTo Erro_Traz_OriDest_Tela

    'move a chave primaria da tabela de preco.. pois ela
    'contem o codigo da origem e destino..
    objTabPreco.lCodigo = objSolServ.lCodTabPreco
    
    'le a tabela de preco adequada...
    lErro = CF("TabPreco_Le2", objTabPreco, objSolServ)
    If lErro <> SUCESSO And lErro <> 96771 Then gError 98613
    
    If lErro <> SUCESSO Then gError 98614
    
    'Ja traz para tela os campos referentes a tabela de preco
    If Len(Trim(LabelAdValoren.Caption)) = 0 Then
        LabelAdValoren.Caption = objTabPreco.dAdValoren * 100 & "%"
    End If
    
    If Len(Trim(LabelPedagio.Caption)) = 0 Then
        LabelPedagio.Caption = Format(objTabPreco.dPedagio, "standard")
    End If
    
    'Coloca em objOriDest o codigo da Origem
    objOriDest.iCodigo = objTabPreco.iOrigem
    
    'Le a Origem
    lErro = CF("OrigemDestino_Le", objOriDest)
    If lErro <> SUCESSO And lErro <> 96567 Then gError 97488
    
    If lErro <> SUCESSO Then gError 98586
    
    'Coloca na Tela os dados da Origem
    LabelOrigem.Caption = objOriDest.sOrigemDestino
    LabelUFOrigem.Caption = objOriDest.sUF
    
    'Coloca em objOriDest o codigo do destino
    objOriDest.iCodigo = objTabPreco.iDestino
    
    'Le o destino
    lErro = CF("OrigemDestino_Le", objOriDest)
    If lErro <> SUCESSO And lErro <> 96567 Then gError 97489
    
    If lErro <> SUCESSO Then gError 98594
    
    'coloca na tela os dados do Destino
    LabelDestino.Caption = objOriDest.sOrigemDestino
    LabelUFDestino.Caption = objOriDest.sUF

    Traz_OriDest_Tela = SUCESSO
    
    Exit Function

Erro_Traz_OriDest_Tela:

    Traz_OriDest_Tela = gErr
    
    Select Case gErr
    
        Case 97488, 97489
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "OrigemDestino")
        
        Case 98586, 98594
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_TABPRECO_ORIGEMDESTINO", gErr, objTabPreco.lCodigo, objOriDest.iCodigo)
        
        Case 98613
                
        Case 98614
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_TABPRECO_SOLSERV", gErr, objTabPreco.lCodigo, objSolServ.lNumero, objSolServ.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Traz_SolicitacaoServico_Tela(objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Move os dados da tabela SolicitacaoServico para a tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objCompServ As New ClassCompServ
Dim objTabPreco As New ClassTabPreco
Dim dPrecoProd As Double
Dim iQtdProd As Integer

On Error GoTo Erro_Traz_SolicitacaoServico_Tela
     
    'Le os dados da tabela Solicitação de Serviço
    lErro = CF("SolicitacaoServico_Le", objSolicitacaoServico)
    If lErro <> SUCESSO And lErro <> 98085 Then gError 97422

    'Se não encontrar --> Erro
    If lErro = 98085 Then gError 97423
    
    'move os dados da solicitacao para a tela
    LabelMaterial.Caption = objSolicitacaoServico.sMaterial
    
    'se qtdmaterial nao esta preenchida..
    If Len(Trim(MaskQuantMaterial.Text)) = 0 Then
        MaskQuantMaterial.Text = Formata_Estoque(objSolicitacaoServico.dQuantMaterial)
        
    End If
        
    'se valormerc nao esta preenchido..
    If Len(Trim(MaskValorMerc.Text)) = 0 Then
        MaskValorMerc.Text = Format(objSolicitacaoServico.dValorMercadoria, "standard")
    
    End If
        
    'Coloca o Nome Reduzido do cliente no label Cliente
    LabelCliente.Caption = objSolicitacaoServico.sClienteNomeRed
    
    'Coloca o Nome Reduzido do despachante no label Despachante
    LabelDespachante.Caption = objSolicitacaoServico.sDespachanteNomeRed
    
    'se UM nao esta preenchida
    If Len(Trim(TextUM.Text)) = 0 Then
        TextUM.Text = objSolicitacaoServico.sUM
    End If
        
    If objSolicitacaoServico.iTipoEmbalagem <> 0 Then
        'traz o tipo de embalagem pra tela
        lErro = Traz_TipoEmb_Tela(objSolicitacaoServico)
        If lErro <> SUCESSO Then gError 98578
    End If
    
    'alteracao por tulio 23/07/02
    labelPorto.Caption = objSolicitacaoServico.sPorto
    labelBooking.Caption = objSolicitacaoServico.sBooking
               
    'traz o tipocontainer pra tela
    lErro = Traz_TipoContainer_Tela(objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98582
         
    'traz o prognavio pra tela
    lErro = Traz_ProgNavio_Tela(objSolicitacaoServico)
    If lErro <> SUCESSO And lErro <> 97421 Then gError 98584
           
    'traz a origem e o destino
    lErro = Traz_OriDest_Tela(objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98587
    
    MaskSolicitacao.Text = objSolicitacaoServico.lNumero
    
    Traz_SolicitacaoServico_Tela = SUCESSO

    Exit Function

Erro_Traz_SolicitacaoServico_Tela:

    Traz_SolicitacaoServico_Tela = gErr

    Select Case gErr

        Case 97422, 98578, 98582, 98584, 98587
        
        Case 97423
            lErro = Rotina_Erro(vbYesNo, "ERRO_SOLICITACAO_NAO_CADASTRADA1", objSolicitacaoServico.lNumero, objSolicitacaoServico.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Traz_CompServ_Tela(objCompServ As ClassCompServ)
'Move os dados da tabela CompServGR para a tela

Dim lErro As Long
Dim objServico As New ClassServico

On Error GoTo Erro_Traz_CompServ_Tela
                
    Call Limpa_Tela_CompServ
       
    'Le os dados da tabela Comprovante de Servico
    lErro = CF("CompServGR_Le", objCompServ)
    
    If lErro <> SUCESSO And lErro <> 97419 Then gError 97465
    
    'Se não encontrar --> Erro
    If lErro = 97419 Then gError 97466
    
    objServico.sProduto = objCompServ.sProduto
    
    'Le os dados dos itens de serviço
    'relacionados com o servico que se relaciona com
    'o Comp Serv em questão
    lErro = CF("CompServGR_Le_CompServItem", objCompServ)
    If lErro <> SUCESSO Then gError 97467

    'move os dados do compserv para a tela
    lErro = Move_Dados_Tela(objCompServ)
    If lErro <> SUCESSO Then gError 97468
        
    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Traz_CompServ_Tela = SUCESSO

    Exit Function

Erro_Traz_CompServ_Tela:

    Traz_CompServ_Tela = gErr

    Select Case gErr

        Case 97465, 97466, 97467, 97468
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Traz_ProgNavio_Tela(objSolServ As ClassSolicitacaoServico) As Long
'Coloca os dados do código passado como parâmetro na tela

Dim lErro As Long
Dim objProgNavio As New ClassProgNavio

On Error GoTo Erro_Traz_ProgNavio_Tela
            
    'move a chave para o objprognavio
    objProgNavio.lCodigo = objSolServ.lCodProgNavio
            
    'Lê os dados de ProgNavio relacionados ao código passado no objProgNavio
    lErro = CF("ProgNavio_Le", objProgNavio)
    If lErro <> AD_SQL_SUCESSO And lErro <> 96657 Then gError 97420

    'Se não existe o Código passado
    If lErro = 96657 Then gError 97421

    'O Código está cadastrado, coloca os dados do obj na tela
    LabelArmador.Caption = objProgNavio.sArmador
    LabelNavio.Caption = objProgNavio.sNavio
    LabelAgMaritimo.Caption = objProgNavio.sAgMaritima
    LabelViagem.Caption = objProgNavio.sViagem
    LabelProgNavio.Caption = objProgNavio.lCodigo
    
    If objProgNavio.dtHoraChegada <> DATA_NULA Then
        LabelHoraChegada.Caption = objProgNavio.dtHoraChegada
    Else
        LabelHoraChegada.Caption = ""
    End If
    
    If objProgNavio.dtHoraDeadLine <> DATA_NULA Then
        LabelHoraDeadLine.Caption = objProgNavio.dtHoraDeadLine
    Else
        LabelHoraDeadLine.Caption = ""
    End If
    
    If objProgNavio.dtDataChegada <> DATA_NULA Then
        LabelDataChegada.Caption = objProgNavio.dtDataChegada
        MaskDemurrage.Text = Format(objProgNavio.dtDataChegada + DELAY_DEMURRAGE, "dd/mm/yy")
        
    Else
        LabelDataChegada.Caption = ""
    End If
    
    If objProgNavio.dtDataDeadLine <> DATA_NULA Then
        LabelDataDeadLine.Caption = objProgNavio.dtDataDeadLine
    Else
        LabelDataDeadLine.Caption = ""
    End If
    
    Traz_ProgNavio_Tela = SUCESSO

    Exit Function

Erro_Traz_ProgNavio_Tela:

    Traz_ProgNavio_Tela = gErr

    Select Case gErr

        Case 97420
        
        Case 97421
       '     Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_PROGNAVIO_SOLSERV", gErr, objProgNavio.lCodigo, objSolServ.lNumero, objSolServ.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Function Traz_TipoContainer_Tela(objSolServ As ClassSolicitacaoServico) As Long
'Traz os dados do tipocontainer para a tela

Dim lErro As Long
Dim objTipoContainer As New ClassTipoContainer

On Error GoTo Erro_Traz_TipoContainer_Tela

    'copia a chave para o objtipocontainer
    objTipoContainer.iTipo = objSolServ.iTipoContainer
    
    'le o tipocontainer
    lErro = CF("TipoContainer_Le", objTipoContainer)
    If lErro <> SUCESSO And lErro <> 96507 Then gError 98507

    If lErro <> SUCESSO Then gError 98581

    'a descricao nao ira + se repetir..
    LabelTipoContainer.Caption = objTipoContainer.iTipo & SEPARADOR & objTipoContainer.sDescricao
    MaskValorContainer.Text = Format(objTipoContainer.dValor, "standard")
     
    Traz_TipoContainer_Tela = SUCESSO

    Exit Function

Erro_Traz_TipoContainer_Tela:

    Traz_TipoContainer_Tela = gErr
    
    Select Case gErr
    
        Case 98507
    
        Case 98581
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_TIPOCONTAINER_SOLSERV", gErr, objTipoContainer.iTipo, objSolServ.lNumero, objSolServ.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Traz_TipoEmb_Tela(objSolServ As ClassSolicitacaoServico) As Long
'Traz os dados do tipo embalagem relacionado com a solicitacao de
'servico para a tela...

Dim lErro As Long
Dim objTipoEmb As New ClassTipoEmbalagem

On Error GoTo Erro_Traz_TipoEmb_Tela

    'passa a chave para o objtipoemb
    objTipoEmb.iTipo = objSolServ.iTipoEmbalagem
    
    'le o tipo embalagem..
    lErro = CF("TipoEmbalagem_Le", objTipoEmb)
    If lErro <> SUCESSO And lErro <> 96507 Then gError 97486

    If lErro <> SUCESSO Then gError 98579

    'coloca os dados relevantes na tela..
    LabelTipoEmbalagem.Caption = objTipoEmb.iTipo & SEPARADOR & objTipoEmb.sDescricao

    Traz_TipoEmb_Tela = SUCESSO

    Exit Function

Erro_Traz_TipoEmb_Tela:

    Traz_TipoEmb_Tela = gErr
    
    Select Case gErr
    
        Case 97486
    
        Case 98579
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_TIPOEMB_SOLSERV", gErr, objTipoEmb.iTipo, objSolServ.lNumero, objSolServ.iFilialEmpresa)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_ItemServico(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Equipamentos

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Item Serviço")
    objGridInt.colColuna.Add ("Descrição Item")
    objGridInt.colColuna.Add ("Data Previsão")
    objGridInt.colColuna.Add ("Hora Previsão")
    objGridInt.colColuna.Add ("Placa Caminhão")
    objGridInt.colColuna.Add ("Placa Carreta")
    objGridInt.colColuna.Add ("Motorista")
    objGridInt.colColuna.Add ("Data Início")
    objGridInt.colColuna.Add ("Hora Início")
    objGridInt.colColuna.Add ("Data Fim")
    objGridInt.colColuna.Add ("Hora Fim")
    objGridInt.colColuna.Add ("Doc.Int.Tipo")
    objGridInt.colColuna.Add ("Doc.Int.Número")
    objGridInt.colColuna.Add ("Doc.Int.Emissão")
    objGridInt.colColuna.Add ("Doc.Ext.Tipo")
    objGridInt.colColuna.Add ("Doc.Ext.Número")
    objGridInt.colColuna.Add ("Doc.Ext.Emissão")
    objGridInt.colColuna.Add ("Doc.Ext.Recepção")
    objGridInt.colColuna.Add ("Doc.Ext.Hora Rec.")
    objGridInt.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (MaskItemServico.Name)
    objGridInt.colCampo.Add (TextDescItemServico.Name)
    objGridInt.colCampo.Add (MaskDataPrev.Name)
    objGridInt.colCampo.Add (MaskHoraPrev.Name)
    objGridInt.colCampo.Add (MaskPlacaCaminhao.Name)
    objGridInt.colCampo.Add (MaskPlacaCarreta.Name)
    objGridInt.colCampo.Add (MaskMotorista.Name)
    objGridInt.colCampo.Add (MaskDataInicio.Name)
    objGridInt.colCampo.Add (MaskHoraInicio.Name)
    objGridInt.colCampo.Add (MaskDataFim.Name)
    objGridInt.colCampo.Add (MaskHoraFim.Name)
    objGridInt.colCampo.Add (MaskDocIntTipo.Name)
    objGridInt.colCampo.Add (MaskDocIntNumero.Name)
    objGridInt.colCampo.Add (MaskDocIntDataEmi.Name)
    objGridInt.colCampo.Add (MaskDocExtTipo.Name)
    objGridInt.colCampo.Add (MaskDocExtNumero.Name)
    objGridInt.colCampo.Add (MaskDocExtDataEmi.Name)
    objGridInt.colCampo.Add (MaskDocExtDataRec.Name)
    objGridInt.colCampo.Add (MaskDocExtHoraRec.Name)
    objGridInt.colCampo.Add (TextObservacao.Name)
    
    'Colunas do Grid
    iGrid_ItemServico_Col = 1
    iGrid_DescItemServico_Col = 2
    iGrid_DataPrev_Col = 3
    iGrid_HoraPrev_Col = 4
    iGrid_PlacaCaminhao_Col = 5
    iGrid_PlacaCarreta_Col = 6
    iGrid_Motorista_Col = 7
    iGrid_DataInicio_Col = 8
    iGrid_HoraInicio_Col = 9
    iGrid_DataFim_Col = 10
    iGrid_HoraFim_Col = 11
    iGrid_DITipo_Col = 12
    iGrid_DINumero_Col = 13
    iGrid_DIDataEmissao_Col = 14
    iGrid_DETipo_Col = 15
    iGrid_DENumero_Col = 16
    iGrid_DEDataEmissao_Col = 17
    iGrid_DEDataRecepcao_Col = 18
    iGrid_DEHoraRecepcao_Col = 19
    iGrid_Observacao_Col = 20
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItemServico

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_SERVICOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 13

    'Largura da primeira coluna
    GridItemServico.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Proibe a inclusão de linhas do grid por parte do usuario
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_ItemServico = SUCESSO

    Exit Function

End Function

Function Trata_Parametros(Optional objCompServ As ClassCompServ) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma Cotacao foi passada por parametro
    If Not (objCompServ Is Nothing) Then
        
        objCompServ.iFilialEmpresa = giFilialEmpresa
        
        'Traz o comprovante de servico pra tela..
        lErro = Traz_CompServ_Tela(objCompServ)
        If lErro <> SUCESSO And lErro <> 97466 Then gError 97485
        
        If lErro = 97466 Then
            
            'limpar a tela
            Call Limpa_Tela_CompServ
            
            'Colocar o codigo na tela
            MaskCodigo.Text = objCompServ.lCodigo
        
        End If
        
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr

        Case 97485

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub Form_Load()
'Inicializa a tela

Dim lErro As Long
Dim iIndiceFrame As Integer
Dim colEmbalagens As New Collection
Dim colContainers As New Collection
Dim objTipoEmb As ClassTipoEmbalagem
Dim objTipoContainer As ClassTipoContainer

On Error GoTo Erro_Form_Load

    iFrameAtual = TAB_DadosPrincipais
    
    'Mascara o produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MaskServico)
    If lErro <> SUCESSO Then gError 98694
    
    'Torna os frames invisiveis a fim de só tornar visivel o frame correspondente e deixa o primeiro visivel
    Frame1(TAB_DadosPrincipais).Visible = True
    For iIndiceFrame = TAB_Complemento To TAB_EtapasServico
        Frame1(iIndiceFrame).Visible = False
    Next
    
    'Inicializa o Grid Itens de Servico
    lErro = Inicializa_Grid_ItemServico(objGridItemServico)
    If lErro <> SUCESSO Then gError 95401
    
    'Inicializa os Eventos
    Set objEventoCompServ = New AdmEvento
    Set objEventoSolicitacao = New AdmEvento
    Set objEventoServico = New AdmEvento
    Set objEventoDocumentoExt = New AdmEvento
    Set objEventoDocumentoInt = New AdmEvento
    
    
    'Inicializa data com a data atual
    MaskDataEmissao.PromptInclude = False
    MaskDataEmissao.Text = Format(gdtDataHoje, "dd/mm/yy")
    MaskDataEmissao.PromptInclude = True
    
    'Coloca o campo "Quantidade Material" no formato de estoque
    MaskQuantMaterial.Format = FORMATO_ESTOQUE
    
    'Início trecho adicionado por RAfael Menezes em 23/09/2002
    gsServico = ""
    'Fim trecho adicionado por RAfael Menezes em 23/09/2002
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr

        Case 95401, 98694
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
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

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera o comando de setas
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'Caso o usuario queira acessar o browser através da tecla F3.
'Ou gerar o proximo numero automatico usando o F2
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim colSelecao As New Collection
Dim objDoc As New ClassDocumento
Dim sSelecao As String
   
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is MaskSolicitacao Then
            Call LabelSolicitacao_Click
        ElseIf Me.ActiveControl Is MaskCodigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is MaskServico Then
            Call LabelServico_Click(2)
        ElseIf Me.ActiveControl Is MaskDocExtTipo Then

            sSelecao = "TipoDoc = ?"

            colSelecao.Add 1

            'Verifica se o documento foi preenchido
            If Len(Trim(MaskDocExtTipo.ClipText)) > 0 Then objDoc.sNomeReduzido = MaskDocExtTipo.Text
        
            'Chamada do browser
            Call Chama_Tela("DocumentoLista", colSelecao, objDoc, objEventoDocumentoExt, sSelecao)

        ElseIf Me.ActiveControl Is MaskDocIntTipo Then
            
            sSelecao = "TipoDoc = ?"
            
            colSelecao.Add 0
            
            'Verifica se o número do Documento foi preenchido
            If Len(Trim(MaskDocIntTipo.ClipText)) > 0 Then objDoc.sNomeReduzido = MaskDocIntTipo.Text
            
            'Chamada do browser
            Call Chama_Tela("DocumentoLista", colSelecao, objDoc, objEventoDocumentoInt, sSelecao)
   
        End If
        
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
    
        Call BotaoProxNum_Click
    
    End If

End Sub

Private Function CompServ_Codigo_Automatico(lCod As Long) As Long
'funcao que gera o codigo automatico

Dim lErro As Long

On Error GoTo Erro_CompServ_Codigo_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_COMPSERV", "CompServGR", "Codigo", lCod)
    If lErro <> SUCESSO Then gError 97406

    CompServ_Codigo_Automatico = SUCESSO

    Exit Function

Erro_CompServ_Codigo_Automatico:

    CompServ_Codigo_Automatico = gErr

    Select Case gErr

        Case 97406
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Carrega_GridItemServico2(objCompServ As ClassCompServ) As Long
'Preenche o Grid de Itens Servico com o Conteudo
'encontrado na tabela CompServItem
'deve receber um objcompserv carregado

Dim iLinha As Integer
Dim objCompServItem As ClassCompServItem
Dim objItemServ As New ClassItemServico
Dim lErro As Long

On Error GoTo Erro_Carrega_GridItemServico2

    'Limpa o Grid de Servicos
    Call Grid_Limpa(objGridItemServico)
    
    iLinha = 0

    'Preenche o grid com os objetos da coleção de cotacaoservicos
    For Each objCompServItem In objCompServ.colCompServItem
    
       'Atualiza o apontador da linha corrente
       iLinha = iLinha + 1
       
       'move para o objitemserv o codigo do item relacionado com
       'o compserv
       objItemServ.iCodigo = objCompServItem.iCodItemServico
       
       'le o item
       lErro = CF("ItemServico_Le", objItemServ)
       If lErro <> SUCESSO And lErro <> 97035 Then gError 98505

       'Item não está cadastrado
       If lErro = 97035 Then gError 98506
        
       'Coloca no grid os dados do item de servico
       GridItemServico.TextMatrix(iLinha, iGrid_ItemServico_Col) = objItemServ.iCodigo
       GridItemServico.TextMatrix(iLinha, iGrid_DescItemServico_Col) = objItemServ.sDescricao
       
       '????? Formatar as datas --> OK!
       If objCompServItem.dtDataFim <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_DataFim_Col) = Format(objCompServItem.dtDataFim, "dd/mm/yy")
       If objCompServItem.dtDataInicio <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_DataInicio_Col) = Format(objCompServItem.dtDataInicio, "dd/mm/yy")
       If objCompServItem.dtDataPrev <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_DataPrev_Col) = Format(objCompServItem.dtDataPrev, "dd/mm/yy")
       If objCompServItem.dtDocExtDataEmissao <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_DEDataEmissao_Col) = Format(objCompServItem.dtDocExtDataEmissao, "dd/mm/yy")
       If objCompServItem.dtDocExtDataRec <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_DEDataRecepcao_Col) = Format(objCompServItem.dtDocExtDataRec, "dd/mm/yy")
       '????? Formatar a hora --> OK!
       If objCompServItem.dtDocExtHoraRec <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_DEHoraRecepcao_Col) = Format(objCompServItem.dtDocExtHoraRec, "hh:mm:ss")
       
       GridItemServico.TextMatrix(iLinha, iGrid_DENumero_Col) = objCompServItem.sDocExtNumero
       GridItemServico.TextMatrix(iLinha, iGrid_DETipo_Col) = objCompServItem.sDocExtTipo
       
       '????? Formatar a hora --> OK!
       If objCompServItem.dtDocIntDataEmissao <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_DIDataEmissao_Col) = Format(objCompServItem.dtDocIntDataEmissao, "dd/mm/yy")
       
       GridItemServico.TextMatrix(iLinha, iGrid_DINumero_Col) = objCompServItem.sDocIntNumero
       GridItemServico.TextMatrix(iLinha, iGrid_DITipo_Col) = objCompServItem.sDocIntTipo
       
       '????? Formatar as horas --> ok
       If objCompServItem.dtHoraFim <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_HoraFim_Col) = Format(objCompServItem.dtHoraFim, "hh:mm:ss")
       If objCompServItem.dtHoraInicio <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_HoraInicio_Col) = Format(objCompServItem.dtHoraInicio, "hh:mm:ss")
       If objCompServItem.dtHoraPrev <> DATA_NULA Then GridItemServico.TextMatrix(iLinha, iGrid_HoraPrev_Col) = Format(objCompServItem.dtHoraPrev, "hh:mm:ss")
       
       GridItemServico.TextMatrix(iLinha, iGrid_Motorista_Col) = objCompServItem.sMotorista
       GridItemServico.TextMatrix(iLinha, iGrid_Observacao_Col) = objCompServItem.sObservacao
       GridItemServico.TextMatrix(iLinha, iGrid_PlacaCaminhao_Col) = objCompServItem.sPlacaCaminhao
       GridItemServico.TextMatrix(iLinha, iGrid_PlacaCarreta_Col) = objCompServItem.sPlacaCarreta
       
    Next
    
    'atualiza o numero de linhas existentes
    objGridItemServico.iLinhasExistentes = iLinha
    
    Carrega_GridItemServico2 = SUCESSO
    
    Exit Function

Erro_Carrega_GridItemServico2:

    Carrega_GridItemServico2 = gErr

    Select Case gErr

        Case 98505

        Case 98506
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_COMPSERV_ITEMSERV", objCompServ.lCodigo, objCompServ.iFilialEmpresa, objItemServ.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Carrega_GridItemServico(objServico As ClassServico) As Long
'Preenche o Grid de Itens Servico com o Conteudo
'encontrado na tabela servitemserv
'o obj servico nao precisa estar carregado com os itens
'mas precisa armazenar o codigo do servico

Dim iLinha As Integer
Dim objItemServ As New ClassItemServico
Dim objServItemServ As New ClassServItemServ
Dim lErro As Long

On Error GoTo Erro_Carrega_GridItemServico

    'Limpa o Grid de Servicos
    Call Grid_Limpa(objGridItemServico)
    
    iLinha = 0

    'carrega a colecao contida em objservico
    lErro = CF("ServicoItemServico_Le", objServico)
    If lErro <> SUCESSO And lErro <> 97543 Then gError 98557
    
    'Preenche o grid com os objetos da coleção de itens de servico
    For Each objServItemServ In objServico.colServItemServ
    
       iLinha = iLinha + 1
       
       'passa o codigo do item para o objitemserv
       objItemServ.iCodigo = objServItemServ.iCodItemServico
       
       'le o item
       lErro = CF("ItemServico_Le", objItemServ)
       If lErro <> SUCESSO And lErro <> 97035 Then gError 98558

       'Item não está cadastrado
       If lErro = 97035 Then gError 98559
        
       'Coloca no grid os dados do item de servico
       GridItemServico.TextMatrix(iLinha, iGrid_ItemServico_Col) = objItemServ.iCodigo
       GridItemServico.TextMatrix(iLinha, iGrid_DescItemServico_Col) = objItemServ.sDescricao
              
    Next
    
    'atualiza o numero de linhas existentes
    objGridItemServico.iLinhasExistentes = iLinha
    
    Carrega_GridItemServico = SUCESSO
    
    Exit Function

Erro_Carrega_GridItemServico:

    Carrega_GridItemServico = gErr

    Select Case gErr

        Case 98557, 98558

        Case 98559
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEGRIDADE_SERV_ITEMSERV", objServico.sProduto, objItemServ.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_CompServ()
'limpa a tela e sugere a data atual como a de emissao

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CompServ

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
        
        'Labels
        gsServico = ""
        LabelCliente.Caption = ""
        labelDescricao.Caption = ""
        LabelMaterial.Caption = ""
        LabelTipoEmbalagem.Caption = ""
        LabelOrigem.Caption = ""
        LabelUFOrigem.Caption = ""
        LabelDestino.Caption = ""
        LabelUFDestino.Caption = ""
        LabelDespachante.Caption = ""
        LabelTipoContainer.Caption = ""
        LabelProgNavio.Caption = ""
        LabelArmador.Caption = ""
        LabelAgMaritimo.Caption = ""
        LabelNavio.Caption = ""
        LabelViagem.Caption = ""
        LabelDataChegada.Caption = ""
        LabelHoraChegada.Caption = ""
        LabelDataDeadLine.Caption = ""
        LabelHoraDeadLine.Caption = ""
        LabelPrecoUnitario.Caption = ""
        LabelAdValoren.Caption = ""
        LabelPedagio.Caption = ""
        labelBooking.Caption = ""
        labelPorto.Caption = ""
        
        'atualizando data
        MaskDataEmissao.PromptInclude = False
        MaskDataEmissao.Text = Format(gdtDataHoje, "dd/mm/yy")
        MaskDataEmissao.PromptInclude = True
        
        'Combos
        ComboStatus.ListIndex = -1
         
        'Grid
        Call Grid_Limpa(objGridItemServico)
        
     Exit Sub

Erro_Limpa_Tela_CompServ:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objCompServ As New ClassCompServ
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    objCompServ.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objCompServ.lNumIntDoc <> 0 Then

        'Carrega o obj com dados de colcampovalor (so o codigo e a filial)
        objCompServ.lCodigo = colCampoValor.Item("Codigo").vValor
        objCompServ.iFilialEmpresa = giFilialEmpresa
       
        'Traz para tela os dados carregados...
        'O caso de nao encontrar nao foi tratado,
        'pois nesse caso a rotina "Traz_CompServ_Tela"
        'é interrompida e nao traz os dados pra tela
        lErro = Traz_CompServ_Tela(objCompServ)
        If lErro <> SUCESSO Then gError 98515
                                     
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 98515

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCompServ As New ClassCompServ

On Error GoTo Erro_Tela_Extrai

    sTabela = "CompServGR"

    lErro = Move_Tela_Memoria(objCompServ)
    If lErro <> SUCESSO Then gError 98516

    'Preenche a coleção colCampoValor
    colCampoValor.Add "NumIntDoc", objCompServ.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "NumIntDocOrigem", objCompServ.lNumIntDocOrigem, 0, "NumIntDocOrigem"
    colCampoValor.Add "FilialEmpresa", objCompServ.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Codigo", objCompServ.lCodigo, 0, "Codigo"
    colCampoValor.Add "DataEmissao", objCompServ.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "Produto", objCompServ.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "QuantMaterial", objCompServ.dQuantMaterial, 0, "QuantMaterial"
    colCampoValor.Add "UM", objCompServ.sUM, STRING_UM_NOME, "UM"
    colCampoValor.Add "ValorMercadoria", objCompServ.dValorMercadoria, 0, "ValorMercadoria"
    colCampoValor.Add "FretePeso", objCompServ.dFretePeso, 0, "FretePeso"
    colCampoValor.Add "Pedagio", objCompServ.dPedagio, 0, "Pedagio"
    colCampoValor.Add "AdValoren", objCompServ.dAdValoren, 0, "AdValoren"
    colCampoValor.Add "DataDemurrage", objCompServ.dtDataDemurrage, 0, "DataDemurrage"
    colCampoValor.Add "CodigoContainer", objCompServ.sCodigoContainer, STRING_CODCONTAINER, "CodigoContainer"
    colCampoValor.Add "ValorContainer", objCompServ.dValorContainer, 0, "ValorContainer"
    colCampoValor.Add "Tara", objCompServ.dTara, 0, "Tara"
    colCampoValor.Add "Lacre", objCompServ.sLacre, STRING_LACRE, "Lacre"
    colCampoValor.Add "Observacao", objCompServ.sObservacao, STRING_OBSERVACAO_OBSERVACAO, "Observacao"
    colCampoValor.Add "NumIntNota", objCompServ.lNumIntNota, 0, "NumIntNota"
    colCampoValor.Add "Situacao", objCompServ.iSituacao, 0, "Situacao"
    colCampoValor.Add "Quantidade", objCompServ.dQuantidade, 0, "Quantidade"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 98516

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objCompServ As ClassCompServ) As Long
'Move os dados relevantes que estao na tela para a memoria...
'Alguns dados sao ignorados, pois nao serao utilizados na gravacao...

Dim sAux As String, sProduto As String
Dim lErro As Long
Dim objDespachante As New ClassDespachante
Dim objSolServ As New ClassSolicitacaoServico
Dim iProdPreenchido As Integer
Dim lAdValoren As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Move os campos para objcompserv, no formato adequado...
    objCompServ.dAdValoren = PercentParaDbl(LabelAdValoren.Caption)
    objCompServ.dFretePeso = StrParaDbl(LabelPrecoUnitario.Caption)
    objCompServ.dPedagio = StrParaDbl(LabelPedagio.Caption)
    objCompServ.dQuantidade = StrParaDbl(MaskQuantidade.Text)
    objCompServ.dQuantMaterial = StrParaDbl(MaskQuantMaterial.Text)
    objCompServ.dTara = StrParaDbl(MaskTaraContainer.Text)
    objCompServ.dtDataDemurrage = StrParaDate(MaskDemurrage.Text)
    objCompServ.dtDataEmissao = StrParaDate(MaskDataEmissao.Text)
    objCompServ.dValorMercadoria = StrParaDbl(MaskValorMerc.Text)
    
    objCompServ.iFilialEmpresa = giFilialEmpresa
    objCompServ.iSituacao = ComboStatus.ListIndex
    
    objCompServ.lCodigo = StrParaDbl(MaskCodigo.ClipText)
    
    objSolServ.lNumero = StrParaDbl(MaskSolicitacao.ClipText)
    objSolServ.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("SolicitacaoServico_Le", objSolServ)
    If lErro <> SUCESSO And lErro <> 98085 Then gError 98510
    
    'If lErro <> SUCESSO Then gError 98724
        
    objCompServ.lNumIntDocOrigem = objSolServ.lNumIntDoc
    objCompServ.sCodigoContainer = MaskCodigoContainer.ClipText
    objCompServ.dValorContainer = StrParaDbl(MaskValorContainer.ClipText)
    objCompServ.sLacre = TextLacreContainer.Text
    objCompServ.sObservacao = TextObs.Text
    objCompServ.sUM = TextUM.Text
        
    lErro = CF("Produto_Formata", MaskServico.Text, sProduto, iProdPreenchido)
    If lErro <> SUCESSO Then gError 98653
            
    objCompServ.sProduto = sProduto
        
    'Move o conteudo do grid para a memoria...
    lErro = Move_Grid_Memoria(objCompServ)
    If lErro <> SUCESSO Then gError 98511

    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 98510, 98511, 98653
                
'        Case 98724
'            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERV_NAO_EXISTE_MAIS", gErr, objSolServ.lNumero, objSolServ.iFilialEmpresa)
'
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Function Move_Grid_Memoria(objCompServ As ClassCompServ) As Long
'Carrega a colecao contida em objcompserv com os itens
'de servico

Dim iLinha As Integer
Dim objCompServItem As ClassCompServItem
Dim objDoc As New ClassDocumento
Dim lErro As Long

On Error GoTo Erro_Move_Grid_Memoria
   
    For iLinha = 1 To objGridItemServico.iLinhasExistentes
        
        Set objCompServItem = New ClassCompServItem
        
        'Carrega os dados em objCompServItem
        objCompServItem.dtDataFim = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DataFim_Col))
        objCompServItem.dtDataInicio = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DataInicio_Col))
        objCompServItem.dtDataPrev = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DataPrev_Col))
        objCompServItem.dtDocExtDataEmissao = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DEDataEmissao_Col))
        objCompServItem.dtDocExtDataRec = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DEDataRecepcao_Col))
        objCompServItem.dtDocExtHoraRec = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DEHoraRecepcao_Col))
        objCompServItem.dtDocIntDataEmissao = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DIDataEmissao_Col))
        objCompServItem.dtHoraFim = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_HoraFim_Col))
        objCompServItem.dtHoraInicio = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_HoraInicio_Col))
        objCompServItem.dtHoraPrev = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_HoraPrev_Col))
        objCompServItem.iCodItemServico = StrParaInt(GridItemServico.TextMatrix(iLinha, iGrid_ItemServico_Col))
        objCompServItem.sDocExtNumero = GridItemServico.TextMatrix(iLinha, iGrid_DENumero_Col)
        objCompServItem.sDocExtTipo = GridItemServico.TextMatrix(iLinha, iGrid_DETipo_Col)
        objCompServItem.sDocIntNumero = GridItemServico.TextMatrix(iLinha, iGrid_DINumero_Col)
        objCompServItem.sDocIntTipo = GridItemServico.TextMatrix(iLinha, iGrid_DITipo_Col)
        objCompServItem.sMotorista = GridItemServico.TextMatrix(iLinha, iGrid_Motorista_Col)
        objCompServItem.sObservacao = GridItemServico.TextMatrix(iLinha, iGrid_Observacao_Col)
        objCompServItem.sPlacaCaminhao = GridItemServico.TextMatrix(iLinha, iGrid_PlacaCaminhao_Col)
        objCompServItem.sPlacaCarreta = GridItemServico.TextMatrix(iLinha, iGrid_PlacaCarreta_Col)
                              
'        'armazena nos campos string o codigo...
'        'obs. eh feito nos campos string pois nao ha campo inteiro para
'        'armazenar o codigo do documento interno....
'        If Len(Trim(objCompServItem.sDocExtTipo)) <> 0 Then
'
'            objDoc.sNomeReduzido = objCompServItem.sDocExtTipo
'
'            'nessa chamada, assumo que o documento ja existe... pois a saida
'            'celula do controle ja foi executada
'            lErro = Documento_Le_NomeRed(objDoc)
'            If lErro <> SUCESSO And lErro <> 98548 Then gError 98653
'
'            If lErro <> SUCESSO Then gError 98654
'
'            objCompServItem.sDocExtTipo = CStr(objDoc.iCodigo)
'
'        End If
                
               
'        If Len(Trim(objCompServItem.sDocIntTipo)) <> 0 Then
'
'            objDoc.sNomeReduzido = objCompServItem.sDocIntTipo
'
'            'nessa chamada, assumo que o documento ja existe... pois a saida
'            'celula do controle ja foi executada
'            lErro = TP_Doc_Le(objDoc, objCompServItem.sDocIntTipo)
'            If lErro <> SUCESSO And lErro <> 98548 Then gError 98655
'
'            If lErro <> SUCESSO Then gError 98656
'
'            objCompServItem.sDocIntTipo = CStr(objDoc.iCodigo)
'
'        End If
        
        objCompServ.colCompServItem.Add objCompServItem
        
    Next

    Move_Grid_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = gErr

    Select Case gErr
        
        Case 98653, 98655
                
'        Case 98654, 98656
'            Call Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_EXISTE_MAIS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
           
    End Select

    Exit Function
    
End Function

Function Testa_Consistencia_Tela() As Long
'Testa a consistencia entre datas e horas da tela
'??? colocar a linha do grid depois...
Dim iLinha As Integer
    
On Error GoTo Erro_Testa_Consistencia_Tela
    
    '???? Voou ---> pousei ja....
    
    For iLinha = 1 To objGridItemServico.iLinhasExistentes
    
        'Se horaprev esta preenchida, verifica se dataprev tbm esta..
        If Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_HoraPrev_Col))) > 0 And Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_DataPrev_Col))) = 0 Then gError 98710
    
        'se horainicio esta preenchida, verifica se datainicio tbm esta..
        If Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_HoraInicio_Col))) > 0 And Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_DataInicio_Col))) = 0 Then gError 98711
        
        'se horafim esta preenchida, verifica se datafim tbm esta..
        If Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_HoraFim_Col))) > 0 And Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_DataFim_Col))) = 0 Then gError 98712
        
        'se data inicio e data fim estiverem preenchidos
        If Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_DataInicio_Col))) > 0 And Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_DataFim_Col))) > 0 Then
            'se dataini > que datafim
            If StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DataInicio_Col)) > StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DataFim_Col)) Then
                
                gError 98714
                
            'se for =, se horaini > que horafim, erro
            ElseIf StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DataInicio_Col)) = StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DataFim_Col)) Then
                'se horaini e horafim estiverem preenchidos
                If Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_HoraInicio_Col))) > 0 And Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_HoraFim_Col))) > 0 Then
                    'se horaini > que hora fim, erro
                    If StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_HoraInicio_Col)) > StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_HoraFim_Col)) Then gError 98715
                End If
            End If
         End If
                          
        'se data de recepcao e emissao do documento externo estiverem preenchidas
        If Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_DEDataEmissao_Col))) > 0 And Len(Trim(GridItemServico.TextMatrix(iLinha, iGrid_DEDataRecepcao_Col))) > 0 Then
           If StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DEDataEmissao_Col)) > StrParaDate(GridItemServico.TextMatrix(iLinha, iGrid_DEDataRecepcao_Col)) Then gError 98718
        End If
    
    Next
        
    Testa_Consistencia_Tela = SUCESSO
        
    Exit Function
    
Erro_Testa_Consistencia_Tela:

    Testa_Consistencia_Tela = gErr

    Select Case gErr
    
        Case 98710
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAPREV_SEM_DATAPREV", gErr)
        
        Case 98711
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAINI_SEM_DATAINI", gErr)
    
        Case 98712
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAFIM_SEM_DATAFIM", gErr)
        
        Case 98714
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_MAIOR_DATAFIM", gErr)
            
        Case 98715
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAINI_MAIOR_HORAFIM", gErr)
            
        Case 98718
            Call Rotina_Erro(vbOKOnly, "ERRO_DOCINT_DATAEMI_MAIOR_DATAREC", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select
    
    Exit Function

End Function


VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl TabTributacao 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9075
   ScaleHeight     =   4695
   ScaleWidth      =   9075
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4635
      Index           =   7
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame FrameTributacao 
         BorderStyle     =   0  'None
         Caption         =   "Resumo"
         Height          =   4110
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   450
         Width           =   8700
         Begin VB.Frame Frame1 
            Caption         =   "ICMS"
            Height          =   1635
            Index           =   203
            Left            =   60
            TabIndex        =   48
            Top             =   870
            Width           =   3600
            Begin VB.Frame Frame1 
               Caption         =   "Substituição"
               Height          =   780
               Index           =   202
               Left            =   165
               TabIndex        =   49
               Top             =   720
               Width           =   3255
               Begin VB.Label ICMSSubstValor 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   53
                  Top             =   375
                  Width           =   1080
               End
               Begin VB.Label ICMSSubstBase 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   375
                  TabIndex        =   52
                  Top             =   375
                  Width           =   1080
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
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
                  Index           =   195
                  Left            =   390
                  TabIndex        =   51
                  Top             =   180
                  Width           =   450
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
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
                  Index           =   21
                  Left            =   1740
                  TabIndex        =   50
                  Top             =   180
                  Width           =   450
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crédito"
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
               Index           =   23
               Left            =   1935
               TabIndex        =   59
               Top             =   195
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label ICMSCredito 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   1920
               TabIndex        =   58
               Top             =   405
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label ICMSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1920
               TabIndex        =   57
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label ICMSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   525
               TabIndex        =   56
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base"
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
               Index           =   208
               Left            =   555
               TabIndex        =   55
               Top             =   165
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
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
               Index           =   22
               Left            =   1920
               TabIndex        =   54
               Top             =   165
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "IPI"
            Height          =   1620
            Index           =   207
            Left            =   3795
            TabIndex        =   41
            Top             =   870
            Width           =   2124
            Begin VB.Label IPICredito 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   855
               TabIndex        =   46
               Top             =   900
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label IPIBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   870
               TabIndex        =   45
               Top             =   375
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   27
               Left            =   300
               TabIndex        =   44
               Top             =   465
               Width           =   495
            End
            Begin VB.Label IPIValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   855
               TabIndex        =   43
               Top             =   900
               Width           =   1080
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
               Height          =   195
               Index           =   28
               Left            =   285
               TabIndex        =   42
               Top             =   930
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crédito:"
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
               Index           =   24
               Left            =   990
               TabIndex        =   47
               Top             =   990
               Visible         =   0   'False
               Width           =   660
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "CSLL"
            Height          =   570
            Index           =   208
            Left            =   4560
            TabIndex        =   38
            Top             =   3510
            Width           =   1860
            Begin MSMask.MaskEdBox CSLLRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   39
               Top             =   195
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Left            =   75
               TabIndex        =   40
               Top             =   270
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "COFINS"
            Height          =   1020
            Index           =   17
            Left            =   6525
            TabIndex        =   33
            Top             =   2535
            Width           =   1860
            Begin MSMask.MaskEdBox COFINSRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   34
               Top             =   630
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox COFINSValor 
               Height          =   285
               Left            =   750
               TabIndex        =   35
               Top             =   210
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Left            =   75
               TabIndex        =   37
               Top             =   705
               Width           =   630
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
               Height          =   195
               Index           =   16
               Left            =   195
               TabIndex        =   36
               Top             =   255
               Width           =   510
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "IR"
            Height          =   1356
            Index           =   19
            Left            =   2670
            TabIndex        =   26
            Top             =   2550
            Width           =   1812
            Begin MSMask.MaskEdBox IRAliquota 
               Height          =   285
               Left            =   600
               TabIndex        =   27
               Top             =   600
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorIRRF 
               Height          =   285
               Left            =   600
               TabIndex        =   28
               Top             =   975
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label IRBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   600
               TabIndex        =   32
               Top             =   240
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   29
               Left            =   75
               TabIndex        =   31
               Top             =   285
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%:"
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
               Index           =   34
               Left            =   276
               TabIndex        =   30
               Top             =   684
               Width           =   216
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
               Height          =   195
               Index           =   35
               Left            =   75
               TabIndex        =   29
               Top             =   1035
               Width           =   510
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "INSS"
            Height          =   1485
            Index           =   21
            Left            =   105
            TabIndex        =   18
            Top             =   2550
            Width           =   2490
            Begin VB.CheckBox INSSRetido 
               Caption         =   "Retido"
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
               Left            =   1155
               TabIndex        =   19
               Top             =   1170
               Width           =   930
            End
            Begin MSMask.MaskEdBox INSSValor 
               Height          =   285
               Left            =   1140
               TabIndex        =   20
               Top             =   885
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox INSSBase 
               Height          =   285
               Left            =   1140
               TabIndex        =   21
               Top             =   210
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox INSSDeducoes 
               Height          =   285
               Left            =   1140
               TabIndex        =   22
               Top             =   555
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
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
               Height          =   195
               Index           =   38
               Left            =   570
               TabIndex        =   25
               Top             =   945
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   44
               Left            =   570
               TabIndex        =   24
               Top             =   255
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Deduções:"
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
               Index           =   48
               Left            =   150
               TabIndex        =   23
               Top             =   600
               Width           =   930
            End
         End
         Begin VB.CommandButton TributacaoRecalcular 
            Caption         =   "Recalcular Tributação"
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
            Left            =   6540
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   3570
            Width           =   1815
         End
         Begin VB.Frame Frame1 
            Caption         =   "PIS"
            Height          =   1005
            Index           =   22
            Left            =   4560
            TabIndex        =   12
            Top             =   2520
            Width           =   1860
            Begin MSMask.MaskEdBox PISRetido 
               Height          =   285
               Left            =   780
               TabIndex        =   13
               Top             =   600
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PISValor 
               Height          =   285
               Left            =   780
               TabIndex        =   14
               Top             =   195
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   49
               Left            =   105
               TabIndex        =   16
               Top             =   675
               Width           =   630
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
               Height          =   195
               Index           =   50
               Left            =   195
               TabIndex        =   15
               Top             =   225
               Width           =   510
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "ISS"
            Height          =   1695
            Index           =   24
            Left            =   6030
            TabIndex        =   2
            Top             =   840
            Width           =   2445
            Begin VB.CheckBox ISSIncluso 
               Caption         =   "Incluso"
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
               Left            =   1290
               TabIndex        =   3
               Top             =   645
               Width           =   975
            End
            Begin MSMask.MaskEdBox ISSAliquota 
               Height          =   285
               Left            =   765
               TabIndex        =   4
               Top             =   600
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ISSValor 
               Height          =   285
               Left            =   750
               TabIndex        =   5
               Top             =   960
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ISSRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   6
               Top             =   1320
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   31
               Left            =   75
               TabIndex        =   11
               Top             =   1395
               Width           =   630
            End
            Begin VB.Label ISSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   750
               TabIndex        =   10
               Top             =   210
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   32
               Left            =   240
               TabIndex        =   9
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%:"
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
               Index           =   33
               Left            =   510
               TabIndex        =   8
               Top             =   645
               Width           =   210
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
               Height          =   195
               Index           =   54
               Left            =   225
               TabIndex        =   7
               Top             =   1005
               Width           =   510
            End
         End
         Begin MSMask.MaskEdBox TipoTributacao 
            Height          =   330
            Left            =   2055
            TabIndex        =   60
            Top             =   510
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label NatOpInternaEspelho 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2055
            TabIndex        =   125
            Top             =   90
            Width           =   525
         End
         Begin VB.Label DescNatOpInterna 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2670
            TabIndex        =   124
            Top             =   75
            Width           =   5610
         End
         Begin VB.Label LblNatOpInternaEspelho 
            AutoSize        =   -1  'True
            Caption         =   "Natureza de Oper.:"
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
            TabIndex        =   123
            Top             =   150
            Width           =   1575
         End
         Begin VB.Label LblTipoTrib 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Tributação:"
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
            Left            =   330
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   62
            Top             =   585
            Width           =   1695
         End
         Begin VB.Label DescTipoTrib 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2685
            TabIndex        =   61
            Top             =   510
            Width           =   5610
         End
      End
      Begin VB.Frame FrameTributacao 
         BorderStyle     =   0  'None
         Caption         =   "Detalhamento"
         Height          =   4110
         Index           =   2
         Left            =   120
         TabIndex        =   63
         Top             =   420
         Visible         =   0   'False
         Width           =   8700
         Begin VB.Frame Frame1 
            Caption         =   "Sobre"
            Height          =   1275
            Index           =   201
            Left            =   135
            TabIndex        =   100
            Top             =   15
            Width           =   8490
            Begin VB.Frame FrameItensTrib 
               Caption         =   "Item"
               Height          =   645
               Left            =   156
               TabIndex        =   115
               Top             =   528
               Width           =   8190
               Begin VB.ComboBox ComboItensTrib 
                  Height          =   315
                  Left            =   165
                  Style           =   2  'Dropdown List
                  TabIndex        =   116
                  Top             =   228
                  Width           =   3180
               End
               Begin VB.Label LabelUMItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   7335
                  TabIndex        =   121
                  Top             =   240
                  Width           =   765
               End
               Begin VB.Label LabelQtdeItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   6450
                  TabIndex        =   120
                  Top             =   240
                  Width           =   840
               End
               Begin VB.Label LabelValorItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   4050
                  TabIndex        =   119
                  Top             =   270
                  Width           =   1140
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   191
                  Left            =   3465
                  TabIndex        =   118
                  Top             =   300
                  Width           =   570
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   9
                  Left            =   5295
                  TabIndex        =   117
                  Top             =   285
                  Width           =   1065
               End
            End
            Begin VB.Frame FrameOutrosTrib 
               Height          =   645
               Left            =   90
               TabIndex        =   106
               Top             =   525
               Visible         =   0   'False
               Width           =   8250
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Outras Desp.:"
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
                  Index           =   10
                  Left            =   3690
                  TabIndex        =   114
                  Top             =   300
                  Width           =   1185
               End
               Begin VB.Label LabelValorOutrasDespesas 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   4905
                  TabIndex        =   113
                  Top             =   285
                  Width           =   1140
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Seguro:"
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
                  Index           =   192
                  Left            =   1785
                  TabIndex        =   112
                  Top             =   300
                  Width           =   675
               End
               Begin VB.Label LabelValorSeguro 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   2490
                  TabIndex        =   111
                  Top             =   285
                  Width           =   1140
               End
               Begin VB.Label Label1 
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
                  Index           =   2
                  Left            =   6090
                  TabIndex        =   110
                  Top             =   315
                  Width           =   885
               End
               Begin VB.Label LabelValorDesconto 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   6990
                  TabIndex        =   109
                  Top             =   300
                  Width           =   1140
               End
               Begin VB.Label Label1 
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
                  Index           =   25
                  Left            =   60
                  TabIndex        =   108
                  Top             =   300
                  Width           =   510
               End
               Begin VB.Label LabelValorFrete 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   600
                  TabIndex        =   107
                  Top             =   285
                  Width           =   1140
               End
            End
            Begin VB.OptionButton TribSobreOutrasDesp 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   4623
               TabIndex        =   105
               Top             =   225
               Width           =   1845
            End
            Begin VB.OptionButton TribSobreSeguro 
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
               Height          =   288
               Left            =   2978
               TabIndex        =   104
               Top             =   225
               Width           =   960
            End
            Begin VB.OptionButton TribSobreDesconto 
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
               Height          =   288
               Left            =   7155
               TabIndex        =   103
               Top             =   225
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.OptionButton TribSobreFrete 
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
               Height          =   288
               Left            =   1477
               TabIndex        =   102
               Top             =   225
               Width           =   816
            End
            Begin VB.OptionButton TribSobreItem 
               Caption         =   "Item"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   108
               TabIndex        =   101
               Top             =   225
               Width           =   684
            End
         End
         Begin VB.Frame Frame1 
            Height          =   2745
            Index           =   211
            Left            =   135
            TabIndex        =   64
            Top             =   1305
            Width           =   8508
            Begin VB.Frame Frame1 
               Caption         =   "IPI"
               Height          =   2472
               Index           =   18
               Left            =   6000
               TabIndex        =   83
               Top             =   180
               Width           =   2376
               Begin VB.CheckBox IPICredita 
                  Caption         =   "Debita"
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
                  Height          =   264
                  Left            =   1008
                  TabIndex        =   85
                  Top             =   2160
                  Width           =   936
               End
               Begin VB.ComboBox ComboIPITipo 
                  Height          =   315
                  Left            =   270
                  Style           =   2  'Dropdown List
                  TabIndex        =   84
                  Top             =   240
                  Width           =   1716
               End
               Begin MSMask.MaskEdBox IPIPercRedBaseItem 
                  Height          =   288
                  Left            =   1272
                  TabIndex        =   86
                  Top             =   1032
                  Width           =   696
                  _ExtentX        =   1217
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIValorItem 
                  Height          =   288
                  Left            =   840
                  TabIndex        =   87
                  Top             =   1836
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.0000"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIAliquotaItem 
                  Height          =   288
                  Left            =   852
                  TabIndex        =   88
                  Top             =   1452
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIBaseItem 
                  Height          =   288
                  Left            =   852
                  TabIndex        =   89
                  Top             =   636
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Red. Base"
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
                  Index           =   171
                  Left            =   276
                  TabIndex        =   93
                  Top             =   1104
                  Width           =   888
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Aliq."
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
                  Index           =   209
                  Left            =   276
                  TabIndex        =   92
                  Top             =   1500
                  Width           =   384
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
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
                  Index           =   40
                  Left            =   240
                  TabIndex        =   91
                  Top             =   1896
                  Width           =   456
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
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
                  Index           =   43
                  Left            =   288
                  TabIndex        =   90
                  Top             =   732
                  Width           =   444
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "ICMS"
               Height          =   1692
               Index           =   205
               Left            =   132
               TabIndex        =   65
               Top             =   960
               Width           =   5688
               Begin VB.CheckBox ICMSCredita 
                  Caption         =   "Debita"
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
                  Height          =   264
                  Left            =   2490
                  TabIndex        =   74
                  Top             =   1368
                  Width           =   936
               End
               Begin VB.Frame Frame1 
                  Caption         =   "Substituição"
                  Height          =   1368
                  Index           =   200
                  Left            =   3552
                  TabIndex        =   67
                  Top             =   144
                  Width           =   2004
                  Begin MSMask.MaskEdBox ICMSSubstValorItem 
                     Height          =   288
                     Left            =   672
                     TabIndex        =   68
                     Top             =   984
                     Width           =   1116
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#,##0.0000"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox ICMSSubstAliquotaItem 
                     Height          =   285
                     Left            =   690
                     TabIndex        =   69
                     Top             =   630
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#0.#0\%"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox ICMSSubstBaseItem 
                     Height          =   288
                     Left            =   684
                     TabIndex        =   70
                     Top             =   252
                     Width           =   1092
                     _ExtentX        =   1905
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#,##0.00"
                     PromptChar      =   " "
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Valor"
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
                     Index           =   193
                     Left            =   108
                     TabIndex        =   73
                     Top             =   1020
                     Width           =   456
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Aliq."
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
                     Index           =   181
                     Left            =   180
                     TabIndex        =   72
                     Top             =   672
                     Width           =   384
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Base"
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
                     Index           =   30
                     Left            =   135
                     TabIndex        =   71
                     Top             =   315
                     Width           =   450
                  End
               End
               Begin VB.ComboBox ComboICMSTipo 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   66
                  Top             =   228
                  Width           =   3336
               End
               Begin MSMask.MaskEdBox ICMSValorItem 
                  Height          =   285
                  Left            =   2295
                  TabIndex        =   75
                  Top             =   1005
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.0000"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSAliquotaItem 
                  Height          =   285
                  Left            =   2325
                  TabIndex        =   76
                  Top             =   630
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSPercRedBaseItem 
                  Height          =   288
                  Left            =   1032
                  TabIndex        =   77
                  Top             =   1008
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSBaseItem 
                  Height          =   288
                  Left            =   588
                  TabIndex        =   78
                  Top             =   624
                  Width           =   1116
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Red. Base"
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
                  Index           =   42
                  Left            =   96
                  TabIndex        =   82
                  Top             =   1068
                  Width           =   888
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Aliq."
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
                  Index           =   196
                  Left            =   1812
                  TabIndex        =   81
                  Top             =   648
                  Width           =   384
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
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
                  Index           =   47
                  Left            =   1788
                  TabIndex        =   80
                  Top             =   1044
                  Width           =   456
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
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
                  Index           =   206
                  Left            =   84
                  TabIndex        =   79
                  Top             =   648
                  Width           =   444
               End
            End
            Begin MSMask.MaskEdBox NaturezaOpItem 
               Height          =   300
               Left            =   1530
               TabIndex        =   94
               Top             =   225
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   529
               _Version        =   393216
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
            Begin MSMask.MaskEdBox TipoTributacaoItem 
               Height          =   300
               Left            =   1875
               TabIndex        =   95
               Top             =   645
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin VB.Label DescTipoTribItem 
               BorderStyle     =   1  'Fixed Single
               Height          =   288
               Left            =   2460
               TabIndex        =   99
               Top             =   636
               Width           =   3120
            End
            Begin VB.Label LabelDescrNatOpItem 
               BorderStyle     =   1  'Fixed Single
               Height          =   288
               Left            =   2184
               TabIndex        =   98
               Top             =   228
               Width           =   3432
            End
            Begin VB.Label NaturezaItemLabel 
               AutoSize        =   -1  'True
               Caption         =   "Natureza Oper.:"
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
               Left            =   168
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   97
               Top             =   252
               Width           =   1368
            End
            Begin VB.Label LblTipoTribItem 
               Caption         =   "Tipo de Tributação:"
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
               Height          =   225
               Left            =   90
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   96
               Top             =   675
               Width           =   1710
            End
         End
      End
      Begin MSComctlLib.TabStrip OpcaoTributacao 
         Height          =   4500
         Left            =   105
         TabIndex        =   122
         Top             =   105
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   7938
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Resumo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhamento"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "TabTributacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iValorIRRFAlterado As Integer
Public iISSRetidoAlterado As Integer
Public iPISRetidoAlterado As Integer
Public iPISValorAlterado As Integer
Public iCOFINSRetidoAlterado As Integer
Public iCOFINSValorAlterado As Integer
Public iCSLLRetidoAlterado As Integer

''*** incluidos p/tratamento de tributacao *******************************
Public Property Get gobjTribTab() As ClassTribTab
    gobjTribTab = Parent.gobjTribTab
End Property

Public Property Get iAlterado() As Integer
    iAlterado = Parent.iAlterado
End Property

Public Property Let iAlterado(vData As Integer)
    Parent.iAlterado = iAlterado
End Property

Public Sub NatOpInterna_GotFocus()
    
    Call gobjTribTab.NatOpInterna_GotFocus(iAlterado)

End Sub

Public Sub NaturezaItemLabel_Click()

    Call gobjTribTab.NaturezaItemLabel_Click

End Sub

Public Sub NaturezaOpItem_GotFocus()
    
    Call gobjTribTab.NaturezaOpItem_GotFocus(iAlterado)

End Sub

Public Sub TipoTributacao_GotFocus()
    
    Call gobjTribTab.TipoTributacao_GotFocus(iAlterado)

End Sub

Public Sub TipoTributacaoItem_GotFocus()
    
    Call gobjTribTab.TipoTributacaoItem_GotFocus(iAlterado)

End Sub

Public Sub ValorIRRF_Validate(Cancel As Boolean)

    Call gobjTribTab.ValorIRRF_Validate(Cancel)

End Sub

Public Sub ComboICMSTipo_Click()

    Call gobjTribTab.ComboICMSTipo_Click

End Sub

Public Sub ComboIPITipo_Click()

    Call gobjTribTab.ComboIPITipo_Click

End Sub

Public Sub ComboItensTrib_Click()

    Call gobjTribTab.ComboItensTrib_Click

End Sub

Public Sub LblNatOpInterna_Click()

    Call gobjTribTab.LblNatOpInterna_Click(NF_SAIDA)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set gobjTribTab = Nothing
    
End Sub

Public Sub LblTipoTrib_Click()

    Call gobjTribTab.LblTipoTrib_Click

End Sub

Public Sub LblTipoTribItem_Click()

    Call gobjTribTab.LblTipoTribItem_Click

End Sub

Public Sub NaturezaOpItem_Change()

    Call gobjTribTab.NaturezaOpItem_Change

End Sub

Public Sub NaturezaOpItem_Validate(Cancel As Boolean)

    Call gobjTribTab.NaturezaOpItem_Validate(Cancel)

End Sub

Public Sub TipoTributacao_Change()

    iAlterado = REGISTRO_ALTERADO

    Call gobjTribTab.TipoTributacao_Change

End Sub

Public Sub TipoTributacao_Validate(Cancel As Boolean)

    Call gobjTribTab.TipoTributacao_Validate(Cancel)

End Sub

Public Sub TipoTributacaoItem_Change()

    Call gobjTribTab.TipoTributacaoItem_Change

End Sub

Public Sub TipoTributacaoItem_Validate(Cancel As Boolean)

    Call gobjTribTab.TipoTributacaoItem_Validate(Cancel)

End Sub

Public Sub TribSobreDesconto_Click()

    Call gobjTribTab.TribSobreDesconto_Click

End Sub

Public Sub TribSobreFrete_Click()

    Call gobjTribTab.TribSobreFrete_Click

End Sub

Public Sub TribSobreItem_Click()

    Call gobjTribTab.TribSobreItem_Click

End Sub

Public Sub TribSobreOutrasDesp_Click()

    Call gobjTribTab.TribSobreOutrasDesp_Click

End Sub

Public Sub TribSobreSeguro_Click()

    Call gobjTribTab.TribSobreSeguro_Click

End Sub

Public Sub TributacaoRecalcular_Click()

    Call gobjTribTab.TributacaoRecalcular_Click

End Sub

Public Sub OpcaoTributacao_Click()

    Call gobjTribTab.OpcaoTributacao_Click

End Sub

Public Sub ValorIRRF_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorIRRFAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ValorIRRF_Change

End Sub

Public Sub ICMSAliquotaItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSAliquotaItem_Change

End Sub

Public Sub ICMSAliquotaItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSAliquotaItem_Validate(Cancel)

End Sub

Public Sub ICMSBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSBaseItem_Change

End Sub

Public Sub ICMSBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSBaseItem_Validate(Cancel)

End Sub

Public Sub ICMSPercRedBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSPercRedBaseItem_Change

End Sub

Public Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSPercRedBaseItem_Validate(Cancel)

End Sub

Public Sub ICMSSubstAliquotaItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSubstAliquotaItem_Change

End Sub

Public Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSSubstAliquotaItem_Validate(Cancel)

End Sub

Public Sub ICMSSubstBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSubstBaseItem_Change

End Sub

Public Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSSubstBaseItem_Validate(Cancel)

End Sub

Public Sub ICMSSubstValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSSubstValorItem_Change

End Sub

Public Sub ICMSSubstValorItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSSubstValorItem_Validate(Cancel)

End Sub

Public Sub ICMSValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ICMSValorItem_Change

End Sub

Public Sub ICMSValorItem_Validate(Cancel As Boolean)

    Call gobjTribTab.ICMSValorItem_Validate(Cancel)

End Sub

Public Sub IPIAliquotaItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIAliquotaItem_Change

End Sub

Public Sub IPIAliquotaItem_Validate(Cancel As Boolean)

    Call gobjTribTab.IPIAliquotaItem_Validate(Cancel)

End Sub

Public Sub IPIBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIBaseItem_Change

End Sub

Public Sub IPIBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.IPIBaseItem_Validate(Cancel)

End Sub

Public Sub IPIPercRedBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIPercRedBaseItem_Change

End Sub

Public Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)

    Call gobjTribTab.IPIPercRedBaseItem_Validate(Cancel)

End Sub

Public Sub IPIValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IPIValorItem_Change

End Sub

Public Sub IPIValorItem_Validate(Cancel As Boolean)

    Call gobjTribTab.IPIValorItem_Validate(Cancel)

End Sub

Public Sub IRAliquota_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.IRAliquota_Change

End Sub

Public Sub IRAliquota_Validate(Cancel As Boolean)

    Call gobjTribTab.IRAliquota_Validate(Cancel)

End Sub

Public Sub ISSAliquota_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSAliquota_Change

End Sub

Public Sub ISSAliquota_Validate(Cancel As Boolean)

    Call gobjTribTab.ISSAliquota_Validate(Cancel)

End Sub

Public Sub ISSIncluso_Click()

    Call gobjTribTab.ISSIncluso_Click

End Sub

Public Sub ISSValor_Change()

    Call gobjTribTab.ISSValor_Change

End Sub

Public Sub ISSValor_Validate(Cancel As Boolean)

    Call gobjTribTab.ISSValor_Validate(Cancel)

End Sub
'*** fim tributacao

'jones-15/03/01
Public Sub INSSBase_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.INSSBase_Change
End Sub

Public Sub INSSBase_Validate(Cancel As Boolean)
    Call gobjTribTab.INSSBase_Validate(Cancel)
End Sub

Public Sub INSSDeducoes_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.INSSDeducoes_Change
End Sub

Public Sub INSSDeducoes_Validate(Cancel As Boolean)
    Call gobjTribTab.INSSDeducoes_Validate(Cancel)
End Sub

Public Sub INSSRetido_Click()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.INSSRetido_Click
End Sub

Public Sub INSSValor_Change()
    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.INSSValor_Change
End Sub

Public Sub INSSValor_Validate(Cancel As Boolean)
    Call gobjTribTab.INSSValor_Validate(Cancel)
End Sub
'fim jones-15/03/01

Public Sub PISRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISRetido_Change
    iPISRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PISRetido_Validate(Cancel As Boolean)
    
    If iPISRetidoAlterado = 0 Then Exit Sub
    Call gobjTribTab.PISRetido_Validate(Cancel)
    iPISRetidoAlterado = 0

End Sub

Public Sub ISSRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.ISSRetido_Change
    iISSRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ISSRetido_Validate(Cancel As Boolean)
    
    If iISSRetidoAlterado = 0 Then Exit Sub
    Call gobjTribTab.ISSRetido_Validate(Cancel)
    iISSRetidoAlterado = 0

End Sub

Public Sub COFINSRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSRetido_Change
    iCOFINSRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub COFINSRetido_Validate(Cancel As Boolean)
    
    If iCOFINSRetidoAlterado = 0 Then Exit Sub
    Call gobjTribTab.COFINSRetido_Validate(Cancel)
    iCOFINSRetidoAlterado = 0

End Sub

Public Sub CSLLRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.CSLLRetido_Change
    iCSLLRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CSLLRetido_Validate(Cancel As Boolean)
    
    If iCSLLRetidoAlterado = 0 Then Exit Sub
    Call gobjTribTab.CSLLRetido_Validate(Cancel)
    iCSLLRetidoAlterado = 0

End Sub


Public Sub PISValor_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.PISValor_Change
    iPISValorAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PISValor_Validate(Cancel As Boolean)
    
    If iPISValorAlterado = 0 Then Exit Sub
    Call gobjTribTab.PISValor_Validate(Cancel)
    iPISValorAlterado = 0

End Sub

Public Sub COFINSValor_Change()

    iAlterado = REGISTRO_ALTERADO
    Call gobjTribTab.COFINSValor_Change
    iCOFINSValorAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub COFINSValor_Validate(Cancel As Boolean)
    
    If iCOFINSValorAlterado = 0 Then Exit Sub
    Call gobjTribTab.COFINSValor_Validate(Cancel)
    iCOFINSValorAlterado = 0

End Sub



VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl ContatoClienteOcx 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   10005
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   6165
      Index           =   1
      Left            =   105
      TabIndex        =   92
      Top             =   645
      Width           =   9795
      Begin VB.CheckBox SoAtivos 
         Caption         =   "Só trazer clientes ativos"
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
         Left            =   6210
         TabIndex        =   3
         Top             =   75
         Width           =   2715
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipos de Clientes"
         Height          =   1365
         Left            =   4905
         TabIndex        =   182
         Top             =   4755
         Width           =   4770
         Begin VB.ListBox TiposCliente 
            Height          =   960
            Left            =   135
            Style           =   1  'Checkbox
            TabIndex        =   70
            Top             =   255
            Width           =   4485
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Não Efetuou Compra"
         Height          =   1365
         Index           =   5
         Left            =   90
         TabIndex        =   173
         Top             =   4755
         Width           =   4770
         Begin VB.Frame FrameD5 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   300
            Index           =   0
            Left            =   900
            TabIndex        =   179
            Top             =   225
            Width           =   3630
            Begin MSMask.MaskEdBox DataNCompAte 
               Height          =   300
               Left            =   2235
               TabIndex        =   63
               Top             =   0
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataNCompDe 
               Height          =   300
               Left            =   1350
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataNCompDe 
               Height          =   300
               Left            =   360
               TabIndex        =   61
               Top             =   0
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataNCompAte 
               Height          =   300
               Left            =   3210
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               Index           =   31
               Left            =   0
               TabIndex        =   181
               Top             =   60
               Width           =   315
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
               Height          =   195
               Index           =   30
               Left            =   1830
               TabIndex        =   180
               Top             =   60
               Width           =   360
            End
         End
         Begin VB.Frame FrameD5 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   330
            Index           =   1
            Left            =   1260
            TabIndex        =   177
            Top             =   570
            Width           =   2070
            Begin MSMask.MaskEdBox ApenasDiasNComp 
               Height          =   315
               Left            =   15
               TabIndex        =   66
               Top             =   0
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s) atrás"
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
               Index           =   29
               Left            =   660
               TabIndex        =   178
               Top             =   45
               Width           =   945
            End
         End
         Begin VB.Frame FrameD5 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   360
            Index           =   2
            Left            =   1275
            TabIndex        =   174
            Top             =   915
            Width           =   3435
            Begin MSMask.MaskEdBox EntreDiasNCompDe 
               Height          =   315
               Left            =   0
               TabIndex        =   68
               Top             =   15
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox EntreDiasNCompAte 
               Height          =   315
               Left            =   1860
               TabIndex        =   69
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "e"
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
               Index           =   28
               Left            =   1080
               TabIndex        =   176
               Top             =   60
               Width           =   120
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s) atrás"
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
               Index           =   27
               Left            =   2430
               TabIndex        =   175
               Top             =   60
               Width           =   945
            End
         End
         Begin VB.OptionButton FaixaDataNComp 
            Caption         =   "Faixa"
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
            TabIndex        =   60
            Top             =   270
            Value           =   -1  'True
            Width           =   1620
         End
         Begin VB.OptionButton EntreNComp 
            Caption         =   "Entre"
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
            Left            =   60
            TabIndex        =   67
            Top             =   960
            Width           =   780
         End
         Begin VB.OptionButton ApenasNComp 
            Caption         =   "Desde"
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
            Left            =   75
            TabIndex        =   65
            Top             =   615
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Última Compra"
         Height          =   1725
         Index           =   3
         Left            =   4905
         TabIndex        =   163
         Top             =   3030
         Width           =   4770
         Begin VB.Frame FrameD3 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   1
            Left            =   1095
            TabIndex        =   167
            Top             =   540
            Width           =   3390
            Begin VB.ComboBox ApenasQualifUComp 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":0000
               Left            =   195
               List            =   "ContatoCliente.ctx":000A
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   45
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasuComp 
               Height          =   315
               Left            =   2040
               TabIndex        =   54
               Top             =   45
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   15
               Left            =   2745
               TabIndex        =   168
               Top             =   90
               Width           =   480
            End
         End
         Begin VB.Frame FrameD3 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   735
            Index           =   2
            Left            =   915
            TabIndex        =   169
            Top             =   960
            Width           =   3825
            Begin VB.ComboBox EntreQualifUCompAte 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":0025
               Left            =   2220
               List            =   "ContatoCliente.ctx":0032
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   360
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifUCompDe 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":0050
               Left            =   2220
               List            =   "ContatoCliente.ctx":005D
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Top             =   0
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasUCompDe 
               Height          =   315
               Left            =   375
               TabIndex        =   56
               Top             =   0
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox EntreDiasUCompAte 
               Height          =   315
               Left            =   375
               TabIndex        =   58
               Top             =   360
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "e"
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
               Index           =   18
               Left            =   3615
               TabIndex        =   172
               Top             =   45
               Width           =   120
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   17
               Left            =   1050
               TabIndex        =   171
               Top             =   405
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   16
               Left            =   1050
               TabIndex        =   170
               Top             =   45
               Width           =   480
            End
         End
         Begin VB.Frame FrameD3 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   330
            Index           =   0
            Left            =   840
            TabIndex        =   164
            Top             =   210
            Width           =   3765
            Begin MSMask.MaskEdBox DataUCompAte 
               Height          =   300
               Left            =   2295
               TabIndex        =   50
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataUCompDe 
               Height          =   300
               Left            =   1440
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataUCompDe 
               Height          =   300
               Left            =   450
               TabIndex        =   48
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataUCompAte 
               Height          =   300
               Left            =   3270
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               Index           =   14
               Left            =   90
               TabIndex        =   166
               Top             =   75
               Width           =   315
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
               Height          =   195
               Index           =   13
               Left            =   1890
               TabIndex        =   165
               Top             =   75
               Width           =   360
            End
         End
         Begin VB.OptionButton ApenasUComp 
            Caption         =   "Apenas"
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
            Left            =   60
            TabIndex        =   52
            Top             =   630
            Width           =   1155
         End
         Begin VB.OptionButton EntreUComp 
            Caption         =   "Entre"
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
            Left            =   60
            TabIndex        =   55
            Top             =   990
            Width           =   780
         End
         Begin VB.OptionButton FaixaDataUComp 
            Caption         =   "Faixa"
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
            Left            =   60
            TabIndex        =   47
            Top             =   270
            Value           =   -1  'True
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Primeira Compra"
         Height          =   1725
         Index           =   2
         Left            =   90
         TabIndex        =   153
         Top             =   3030
         Width           =   4770
         Begin VB.OptionButton EntrePComp 
            Caption         =   "Entre"
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
            Left            =   60
            TabIndex        =   42
            Top             =   990
            Width           =   780
         End
         Begin VB.OptionButton ApenasPComp 
            Caption         =   "Apenas"
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
            Left            =   60
            TabIndex        =   39
            Top             =   630
            Width           =   1050
         End
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   330
            Index           =   0
            Left            =   840
            TabIndex        =   160
            Top             =   210
            Width           =   3765
            Begin MSMask.MaskEdBox DataPCompAte 
               Height          =   300
               Left            =   2295
               TabIndex        =   37
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataPCompDe 
               Height          =   300
               Left            =   1440
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataPCompDe 
               Height          =   300
               Left            =   450
               TabIndex        =   35
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataPCompAte 
               Height          =   300
               Left            =   3270
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               Height          =   195
               Index           =   11
               Left            =   1890
               TabIndex        =   162
               Top             =   75
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
               Index           =   9
               Left            =   90
               TabIndex        =   161
               Top             =   75
               Width           =   315
            End
         End
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   1
            Left            =   870
            TabIndex        =   158
            Top             =   540
            Width           =   3825
            Begin VB.ComboBox ApenasQualifPComp 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":007B
               Left            =   420
               List            =   "ContatoCliente.ctx":0085
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   45
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasPComp 
               Height          =   315
               Left            =   2265
               TabIndex        =   41
               Top             =   45
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   8
               Left            =   2970
               TabIndex        =   159
               Top             =   90
               Width           =   480
            End
         End
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   735
            Index           =   2
            Left            =   915
            TabIndex        =   154
            Top             =   960
            Width           =   3825
            Begin VB.ComboBox EntreQualifPCompDe 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":00A0
               Left            =   2220
               List            =   "ContatoCliente.ctx":00AD
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   0
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifPCompAte 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":00CB
               Left            =   2220
               List            =   "ContatoCliente.ctx":00D8
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   360
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasPCompDe 
               Height          =   315
               Left            =   375
               TabIndex        =   43
               Top             =   0
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox EntreDiasPCompAte 
               Height          =   315
               Left            =   375
               TabIndex        =   45
               Top             =   360
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   7
               Left            =   1050
               TabIndex        =   157
               Top             =   45
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Left            =   1050
               TabIndex        =   156
               Top             =   405
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "e"
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
               Index           =   2
               Left            =   3615
               TabIndex        =   155
               Top             =   45
               Width           =   120
            End
         End
         Begin VB.OptionButton FaixaDataPComp 
            Caption         =   "Faixa"
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
            Left            =   60
            TabIndex        =   34
            Top             =   270
            Value           =   -1  'True
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Próximo Contato"
         Height          =   1725
         Index           =   4
         Left            =   90
         TabIndex        =   143
         Top             =   1275
         Width           =   4770
         Begin VB.OptionButton EntrePCont 
            Caption         =   "Entre"
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
            Left            =   60
            TabIndex        =   16
            Top             =   990
            Width           =   780
         End
         Begin VB.OptionButton ApenasPCont 
            Caption         =   "Apenas"
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
            Left            =   60
            TabIndex        =   13
            Top             =   630
            Width           =   1050
         End
         Begin VB.Frame FrameD4 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   330
            Index           =   0
            Left            =   840
            TabIndex        =   150
            Top             =   210
            Width           =   3765
            Begin MSMask.MaskEdBox DataPContAte 
               Height          =   300
               Left            =   2295
               TabIndex        =   11
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataPContDe 
               Height          =   300
               Left            =   1440
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataPContDe 
               Height          =   300
               Left            =   450
               TabIndex        =   9
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataPContAte 
               Height          =   300
               Left            =   3270
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               Height          =   195
               Index           =   25
               Left            =   1890
               TabIndex        =   152
               Top             =   75
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
               Index           =   24
               Left            =   90
               TabIndex        =   151
               Top             =   75
               Width           =   315
            End
         End
         Begin VB.Frame FrameD4 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   1
            Left            =   870
            TabIndex        =   148
            Top             =   540
            Width           =   3825
            Begin VB.ComboBox ApenasQualifPCont 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":00F6
               Left            =   420
               List            =   "ContatoCliente.ctx":0100
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   45
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasPCont 
               Height          =   315
               Left            =   2265
               TabIndex        =   15
               Top             =   45
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   23
               Left            =   2970
               TabIndex        =   149
               Top             =   90
               Width           =   480
            End
         End
         Begin VB.Frame FrameD4 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   735
            Index           =   2
            Left            =   915
            TabIndex        =   144
            Top             =   960
            Width           =   3825
            Begin VB.ComboBox EntreQualifPContDe 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":011B
               Left            =   2220
               List            =   "ContatoCliente.ctx":0128
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   0
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifPContAte 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":0146
               Left            =   2220
               List            =   "ContatoCliente.ctx":0153
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   360
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasPContDe 
               Height          =   315
               Left            =   375
               TabIndex        =   17
               Top             =   0
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox EntreDiasPContAte 
               Height          =   315
               Left            =   375
               TabIndex        =   19
               Top             =   360
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   22
               Left            =   1050
               TabIndex        =   147
               Top             =   45
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   21
               Left            =   1050
               TabIndex        =   146
               Top             =   405
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "e"
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
               Index           =   20
               Left            =   3615
               TabIndex        =   145
               Top             =   45
               Width           =   120
            End
         End
         Begin VB.OptionButton FaixaDataPCont 
            Caption         =   "Faixa"
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
            Left            =   60
            TabIndex        =   8
            Top             =   270
            Value           =   -1  'True
            Width           =   1620
         End
      End
      Begin VB.Frame FrameCategoriaCliente 
         Caption         =   "Categoria de Cliente"
         Height          =   900
         Left            =   90
         TabIndex        =   125
         Top             =   345
         Width           =   9600
         Begin VB.ComboBox CategoriaCliente 
            Height          =   315
            Left            =   3150
            TabIndex        =   5
            Top             =   165
            Width           =   6135
         End
         Begin VB.ComboBox CategoriaClienteDe 
            Height          =   315
            Left            =   1305
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   525
            Width           =   3180
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
            Left            =   315
            TabIndex        =   4
            Top             =   225
            Width           =   855
         End
         Begin VB.ComboBox CategoriaClienteAte 
            Height          =   315
            Left            =   6105
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   510
            Width           =   3180
         End
         Begin VB.Label Label4 
            Caption         =   "Label5"
            Height          =   15
            Left            =   360
            TabIndex        =   129
            Top             =   720
            Width           =   30
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
            Left            =   5700
            TabIndex        =   128
            Top             =   570
            Width           =   360
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
            Left            =   945
            TabIndex        =   127
            Top             =   570
            Width           =   315
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
            Left            =   2235
            TabIndex        =   126
            Top             =   210
            Width           =   855
         End
      End
      Begin VB.ComboBox RespCallCenter 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   60
         Width           =   2820
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Último Contato"
         Height          =   1725
         Index           =   0
         Left            =   4920
         TabIndex        =   105
         Top             =   1275
         Width           =   4770
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   735
            Index           =   2
            Left            =   915
            TabIndex        =   112
            Top             =   960
            Width           =   3825
            Begin VB.ComboBox EntreQualifContAte 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":0171
               Left            =   2220
               List            =   "ContatoCliente.ctx":017E
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   360
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifContDe 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":019C
               Left            =   2220
               List            =   "ContatoCliente.ctx":01A9
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   0
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasContDe 
               Height          =   315
               Left            =   375
               TabIndex        =   30
               Top             =   0
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox EntreDiasContAte 
               Height          =   315
               Left            =   375
               TabIndex        =   32
               Top             =   360
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "e"
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
               Left            =   3615
               TabIndex        =   115
               Top             =   45
               Width           =   120
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   6
               Left            =   1050
               TabIndex        =   114
               Top             =   405
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Left            =   1050
               TabIndex        =   113
               Top             =   45
               Width           =   480
            End
         End
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   1
            Left            =   1050
            TabIndex        =   110
            Top             =   540
            Width           =   3390
            Begin VB.ComboBox ApenasQualifCont 
               Height          =   315
               ItemData        =   "ContatoCliente.ctx":01C7
               Left            =   240
               List            =   "ContatoCliente.ctx":01D1
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   45
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasCont 
               Height          =   315
               Left            =   2085
               TabIndex        =   28
               Top             =   45
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dia(s)"
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
               Index           =   10
               Left            =   2790
               TabIndex        =   111
               Top             =   90
               Width           =   480
            End
         End
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   330
            Index           =   0
            Left            =   840
            TabIndex        =   107
            Top             =   210
            Width           =   3765
            Begin MSMask.MaskEdBox DataContAte 
               Height          =   300
               Left            =   2295
               TabIndex        =   24
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataContDe 
               Height          =   300
               Left            =   1440
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataContDe 
               Height          =   300
               Left            =   450
               TabIndex        =   22
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataContAte 
               Height          =   300
               Left            =   3270
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               Index           =   0
               Left            =   90
               TabIndex        =   109
               Top             =   75
               Width           =   315
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
               Height          =   195
               Index           =   1
               Left            =   1890
               TabIndex        =   108
               Top             =   75
               Width           =   360
            End
         End
         Begin VB.OptionButton FaixaDataCont 
            Caption         =   "Faixa"
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
            Left            =   60
            TabIndex        =   21
            Top             =   270
            Value           =   -1  'True
            Width           =   1620
         End
         Begin VB.OptionButton ApenasCont 
            Caption         =   "Apenas"
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
            Left            =   60
            TabIndex        =   26
            Top             =   630
            Width           =   1050
         End
         Begin VB.OptionButton EntreCont 
            Caption         =   "Entre"
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
            Left            =   60
            TabIndex        =   29
            Top             =   990
            Width           =   780
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Responsável - Call Center:"
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
         Index           =   12
         Left            =   930
         TabIndex        =   106
         Top             =   105
         Width           =   2265
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   6120
      Index           =   2
      Left            =   90
      TabIndex        =   93
      Top             =   645
      Visible         =   0   'False
      Width           =   9795
      Begin VB.Frame Frame1 
         Caption         =   "Clientes"
         Height          =   6060
         Index           =   1
         Left            =   15
         TabIndex        =   94
         Top             =   15
         Width           =   9765
         Begin VB.TextBox DataUContato 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   3840
            TabIndex        =   139
            Top             =   2115
            Width           =   945
         End
         Begin VB.TextBox DataPCompra 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   3930
            TabIndex        =   138
            Top             =   1620
            Width           =   945
         End
         Begin VB.TextBox DataUCompra 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   2865
            TabIndex        =   137
            Top             =   1695
            Width           =   945
         End
         Begin VB.Frame Frame3 
            Caption         =   "Contato"
            Height          =   1455
            Left            =   60
            TabIndex        =   130
            Top             =   4500
            Width           =   9630
            Begin VB.CommandButton BotaoGravarContato 
               Caption         =   "Gravar Contato"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8070
               TabIndex        =   134
               ToolTipText     =   "Gravar Contato"
               Top             =   975
               Width           =   1485
            End
            Begin VB.CommandButton BotaoContatos 
               Caption         =   "Contatos Anteriores"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8070
               TabIndex        =   133
               Top             =   540
               Width           =   1485
            End
            Begin VB.ComboBox Status 
               Height          =   315
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   195
               Width           =   2460
            End
            Begin VB.TextBox Assunto 
               Height          =   855
               Left            =   990
               MaxLength       =   510
               MultiLine       =   -1  'True
               TabIndex        =   79
               Top             =   540
               Width           =   7050
            End
            Begin MSComCtl2.UpDown UpDownDataProx 
               Height          =   300
               Left            =   5715
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   195
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataProx 
               Height          =   300
               Left            =   4755
               TabIndex        =   77
               ToolTipText     =   "Informe a data prevista para o recebimento."
               Top             =   195
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Próx. Contato:"
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
               Left            =   3525
               TabIndex        =   142
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label9 
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
               Left            =   6120
               TabIndex        =   141
               Top             =   225
               Width           =   660
            End
            Begin VB.Label lblCliente 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6840
               TabIndex        =   140
               Top             =   180
               Width           =   2730
            End
            Begin VB.Label Label7 
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
               Height          =   195
               Left            =   330
               TabIndex        =   132
               Top             =   240
               Width           =   615
            End
            Begin VB.Label LabelAssunto 
               AutoSize        =   -1  'True
               Caption         =   "Assunto:"
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
               TabIndex        =   131
               Top             =   540
               Width           =   750
            End
         End
         Begin MSMask.MaskEdBox ValorTotalCompras 
            Height          =   225
            Left            =   510
            TabIndex        =   124
            Top             =   1500
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ligação"
            Height          =   1380
            Left            =   1845
            TabIndex        =   116
            Top             =   3090
            Width           =   5940
            Begin MSMask.MaskEdBox OperDDD 
               Height          =   315
               Left            =   1935
               TabIndex        =   87
               Top             =   975
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   2
               Mask            =   "##"
               PromptChar      =   "_"
            End
            Begin VB.TextBox SenhaTel 
               Enabled         =   0   'False
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   2970
               PasswordChar    =   "*"
               TabIndex        =   86
               Top             =   585
               Width           =   675
            End
            Begin VB.CommandButton BotaoLigar 
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
               Left            =   4620
               Picture         =   "ContatoCliente.ctx":01EC
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   150
               Width           =   405
            End
            Begin VB.CommandButton BotaoDesligar 
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
               Left            =   5085
               Picture         =   "ContatoCliente.ctx":08A6
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   150
               Width           =   405
            End
            Begin VB.ComboBox Contato 
               Height          =   315
               Left            =   1020
               TabIndex        =   82
               ToolTipText     =   $"ContatoCliente.ctx":0F60
               Top             =   210
               Width           =   1785
            End
            Begin VB.CheckBox PossuiSenha 
               Caption         =   "Telefone com senha"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   210
               TabIndex        =   85
               Top             =   585
               Width           =   2055
            End
            Begin MSMask.MaskEdBox DigDiscExt 
               Height          =   315
               Left            =   5190
               TabIndex        =   88
               Top             =   975
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   1
               Mask            =   "#"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               Caption         =   "Para obter linha externa discar:"
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
               Left            =   2490
               TabIndex        =   123
               Top             =   1035
               Width           =   2745
            End
            Begin VB.Label Label5 
               Caption         =   "Operadora de DDD:"
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
               Left            =   210
               TabIndex        =   122
               Top             =   1020
               Width           =   1725
            End
            Begin VB.Label UF 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4065
               TabIndex        =   121
               Top             =   585
               Width           =   495
            End
            Begin VB.Label Label2 
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
               Height          =   240
               Left            =   3720
               TabIndex        =   120
               Top             =   645
               Width           =   315
            End
            Begin VB.Label Label3 
               Caption         =   "Senha:"
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
               Left            =   2370
               TabIndex        =   119
               Top             =   660
               Width           =   660
            End
            Begin VB.Label LabelContato 
               AutoSize        =   -1  'True
               Caption         =   "Contato:"
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
               Left            =   210
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   118
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Telefone 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2805
               TabIndex        =   117
               Top             =   210
               Width           =   1770
            End
         End
         Begin VB.CommandButton BotaoMarcar 
            Caption         =   "Marcar Todas"
            Height          =   555
            Left            =   60
            Picture         =   "ContatoCliente.ctx":0FE8
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   3180
            Width           =   1440
         End
         Begin VB.CommandButton BotaoDesmarcar 
            Caption         =   "Desmarcar Todas"
            Height          =   555
            Left            =   60
            Picture         =   "ContatoCliente.ctx":2002
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   3900
            Width           =   1440
         End
         Begin VB.CommandButton BotaoHistorico 
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
            Height          =   405
            Left            =   8175
            TabIndex        =   89
            Top             =   3180
            Width           =   1485
         End
         Begin VB.CommandButton BotaoGravarTela 
            Caption         =   "Gravar Histórico Dia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   8175
            TabIndex        =   91
            ToolTipText     =   "Gravar Histórico"
            Top             =   4050
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
            Height          =   405
            Left            =   8175
            TabIndex        =   90
            Top             =   3615
            Width           =   1485
         End
         Begin VB.TextBox HistoricoGrid 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   5160
            TabIndex        =   104
            Top             =   1320
            Width           =   3120
         End
         Begin VB.TextBox Fone2Grid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   6705
            TabIndex        =   103
            Top             =   255
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox CheckLigacaoRealizada 
            Height          =   240
            Left            =   7080
            TabIndex        =   102
            Top             =   915
            Width           =   795
         End
         Begin VB.CheckBox CheckLigar 
            Height          =   240
            Left            =   600
            TabIndex        =   101
            Top             =   870
            Width           =   555
         End
         Begin VB.TextBox Fone1Grid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   5550
            TabIndex        =   99
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox ClienteGrid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   2265
            TabIndex        =   98
            Top             =   885
            Width           =   2100
         End
         Begin VB.TextBox FilialClienteGrid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1260
            TabIndex        =   97
            Top             =   1140
            Width           =   765
         End
         Begin MSMask.MaskEdBox ContatoGrid 
            Height          =   240
            Left            =   4365
            TabIndex        =   100
            Top             =   240
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridClientes 
            Height          =   2700
            Left            =   45
            TabIndex        =   75
            Top             =   210
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   4763
            _Version        =   393216
         End
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   525
      Left            =   2670
      TabIndex        =   135
      Top             =   15
      Width           =   5070
      Begin VB.CheckBox OpcaoPadrao 
         Caption         =   "Padrão"
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
         Left            =   3615
         TabIndex        =   1
         Top             =   195
         Width           =   930
      End
      Begin VB.ComboBox OpcoesTela 
         Height          =   315
         Left            =   675
         TabIndex        =   0
         Top             =   135
         Width           =   2820
      End
      Begin VB.Label LabelOpcao 
         AutoSize        =   -1  'True
         Caption         =   "Opção:"
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
         Left            =   30
         TabIndex        =   136
         Top             =   195
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7830
      ScaleHeight     =   450
      ScaleWidth      =   2055
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   45
      Width           =   2115
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "ContatoCliente.ctx":31E4
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "ContatoCliente.ctx":3362
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "ContatoCliente.ctx":3894
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Gravar Opção"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "ContatoCliente.ctx":39EE
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6570
      Left            =   30
      TabIndex        =   96
      Top             =   270
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   11589
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Relacionamentos"
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
   Begin MSCommLib.MSComm ComDiscar 
      Left            =   1020
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      InBufferSize    =   2000
      OutBufferSize   =   2000
      ParityReplace   =   48
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   90
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   555
      Top             =   15
   End
End
Attribute VB_Name = "ContatoClienteOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTContatoCliente
Attribute objCT.VB_VarHelpID = -1

Private Sub RespCallCenter_Validate(Cancel As Boolean)
    objCT.RespCallCenter_Validate (Cancel)
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTContatoCliente
    Set objCT.objUserControl = Me
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
     Call objCT.TabStrip1_BeforeClick(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub GridClientes_Click()
     Call objCT.GridClientes_Click
End Sub

Private Sub GridClientes_GotFocus()
     Call objCT.GridClientes_GotFocus
End Sub

Private Sub GridClientes_EnterCell()
     Call objCT.GridClientes_EnterCell
End Sub

Private Sub GridClientes_LeaveCell()
     Call objCT.GridClientes_LeaveCell
End Sub

Private Sub GridClientes_KeyPress(KeyAscii As Integer)
     Call objCT.GridClientes_KeyPress(KeyAscii)
End Sub

Private Sub GridClientes_RowColChange()
     Call objCT.GridClientes_RowColChange
End Sub

Private Sub GridClientes_Scroll()
     Call objCT.GridClientes_Scroll
End Sub

Private Sub GridClientes_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridClientes_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridClientes_LostFocus()
     Call objCT.GridClientes_LostFocus
End Sub

Private Sub BotaoDesligar_Click()
     Call objCT.BotaoDesligar_Click
End Sub

Private Sub BotaoLigar_Click()
     Call objCT.BotaoLigar_Click
End Sub

Private Sub BotaoGravarTela_Click()
     Call objCT.BotaoGravarTela_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub UpDownDataUCompAte_DownClick()
     Call objCT.UpDownDataUCompAte_DownClick
End Sub

Private Sub UpDownDataUCompAte_UpClick()
     Call objCT.UpDownDataUCompAte_UpClick
End Sub

Private Sub UpDownDataUCompDe_DownClick()
     Call objCT.UpDownDataUCompDe_DownClick
End Sub

Private Sub UpDownDataUCompDe_UpClick()
     Call objCT.UpDownDataUCompDe_UpClick
End Sub

Private Sub UpDownDataContAte_DownClick()
     Call objCT.UpDownDataContAte_DownClick
End Sub

Private Sub UpDownDataContAte_UpClick()
     Call objCT.UpDownDataContAte_UpClick
End Sub

Private Sub UpDownDataContDe_DownClick()
     Call objCT.UpDownDataContDe_DownClick
End Sub

Private Sub UpDownDataContDe_UpClick()
     Call objCT.UpDownDataContDe_UpClick
End Sub

Private Sub UpDownDataPCompAte_DownClick()
     Call objCT.UpDownDataPCompAte_DownClick
End Sub

Private Sub UpDownDataPCompAte_UpClick()
     Call objCT.UpDownDataPCompAte_UpClick
End Sub

Private Sub UpDownDataPCompDe_DownClick()
     Call objCT.UpDownDataPCompDe_DownClick
End Sub

Private Sub UpDownDataPCompDe_UpClick()
     Call objCT.UpDownDataPCompDe_UpClick
End Sub

Private Sub BotaoDesmarcar_Click()
     Call objCT.BotaoDesmarcar_Click
End Sub

Private Sub BotaoMarcar_Click()
     Call objCT.BotaoMarcar_Click
End Sub

Private Sub OpcoesTela_Validate(Cancel As Boolean)
     Call objCT.OpcoesTela_Validate(Cancel)
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub OpcoesTela_Click()
     Call objCT.OpcoesTela_Click
End Sub

Private Sub ApenasCont_Click()
     Call objCT.ApenasCont_Click
End Sub

Private Sub EntreCont_Click()
     Call objCT.EntreCont_Click
End Sub

Private Sub FaixaDataCont_Click()
     Call objCT.FaixaDataCont_Click
End Sub

Private Sub ApenasPComp_Click()
     Call objCT.ApenasPComp_Click
End Sub

Private Sub EntrePComp_Click()
     Call objCT.EntrePComp_Click
End Sub

Private Sub FaixaDataPComp_Click()
     Call objCT.FaixaDataPComp_Click
End Sub

Private Sub ApenasUComp_Click()
     Call objCT.ApenasUComp_Click
End Sub

Private Sub EntreUComp_Click()
     Call objCT.EntreUComp_Click
End Sub

Private Sub FaixaDataUComp_Click()
     Call objCT.FaixaDataUComp_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub DataContDe_Validate(Cancel As Boolean)
     Call objCT.DataContDe_Validate(Cancel)
End Sub

Private Sub DataContAte_Validate(Cancel As Boolean)
     Call objCT.DataContAte_Validate(Cancel)
End Sub

Private Sub DataPCompDe_Validate(Cancel As Boolean)
     Call objCT.DataPCompDe_Validate(Cancel)
End Sub

Private Sub DataPCompAte_Validate(Cancel As Boolean)
     Call objCT.DataPCompAte_Validate(Cancel)
End Sub

Private Sub DataUCompDe_Validate(Cancel As Boolean)
     Call objCT.DataUCompDe_Validate(Cancel)
End Sub

Private Sub DataUCompAte_Validate(Cancel As Boolean)
     Call objCT.DataUCompAte_Validate(Cancel)
End Sub

Private Sub BotaoHistorico_Click()
     Call objCT.BotaoHistorico_Click
End Sub

Private Sub Contato_Click()
     Call objCT.Contato_Click
End Sub

Private Sub Contato_Validate(Cancel As Boolean)
     Call objCT.Contato_Validate(Cancel)
End Sub

Private Sub LabelContato_Click()
     Call objCT.LabelContato_Click
End Sub

Private Sub HistoricoGrid_Change()
     Call objCT.HistoricoGrid_Change
End Sub

Private Sub HistoricoGrid_GotFocus()
     Call objCT.HistoricoGrid_GotFocus
End Sub

Private Sub HistoricoGrid_KeyPress(KeyAscii As Integer)
     Call objCT.HistoricoGrid_KeyPress(KeyAscii)
End Sub

Private Sub HistoricoGrid_Validate(Cancel As Boolean)
     Call objCT.HistoricoGrid_Validate(Cancel)
End Sub

Private Sub CheckLigar_Click()
     Call objCT.CheckLigar_Click
End Sub

Private Sub CheckLigar_GotFocus()
     Call objCT.CheckLigar_GotFocus
End Sub

Private Sub CheckLigar_KeyPress(KeyAscii As Integer)
     Call objCT.CheckLigar_KeyPress(KeyAscii)
End Sub

Private Sub CheckLigar_Validate(Cancel As Boolean)
     Call objCT.CheckLigar_Validate(Cancel)
End Sub

Private Sub CheckLigacaoRealizada_Click()
     Call objCT.CheckLigacaoRealizada_Click
End Sub

Private Sub CheckLigacaoRealizada_GotFocus()
     Call objCT.CheckLigacaoRealizada_GotFocus
End Sub

Private Sub CheckLigacaoRealizada_KeyPress(KeyAscii As Integer)
     Call objCT.CheckLigacaoRealizada_KeyPress(KeyAscii)
End Sub

Private Sub CheckLigacaoRealizada_Validate(Cancel As Boolean)
     Call objCT.CheckLigacaoRealizada_Validate(Cancel)
End Sub

Private Sub BotaoCliente_Click()
     Call objCT.BotaoCliente_Click
End Sub

Private Sub Timer1_Timer()
     Call objCT.Timer1_Timer
End Sub

Private Sub Timer2_Timer()
     Call objCT.Timer2_Timer
End Sub

Private Sub PossuiSenha_Click()
     Call objCT.PossuiSenha_Click
End Sub

Private Sub CategoriaCliente_Click()
     Call objCT.CategoriaCliente_Click
End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)
     Call objCT.CategoriaCliente_Validate(Cancel)
End Sub

Private Sub CategoriaClienteTodas_Click()
     Call objCT.CategoriaClienteTodas_Click
End Sub

Private Sub CategoriaClienteAte_Validate(Cancel As Boolean)
     Call objCT.CategoriaClienteAte_Validate(Cancel)
End Sub

Private Sub CategoriaClienteDe_Validate(Cancel As Boolean)
     Call objCT.CategoriaClienteDe_Validate(Cancel)
End Sub

Private Sub Status_Change()
     Call objCT.Status_Change
End Sub

Private Sub BotaoContatos_Click()
     Call objCT.BotaoContatos_Click
End Sub

Private Sub BotaoGravarContato_Click()
     Call objCT.BotaoGravarContato_Click
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
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

'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call objCT.UserControl_KeyDown(KeyCode, Shift)
'End Sub

Private Sub UpDownDataProx_DownClick()
    Call objCT.UpDownDataProx_DownClick
End Sub

Private Sub UpDownDataProx_UpClick()
    Call objCT.UpDownDataProx_UpClick
End Sub

Private Sub DataProx_Change()
    objCT.DataProx_Change
End Sub

Private Sub DataProx_Validate(Cancel As Boolean)
    Call objCT.DataProx_Validate(Cancel)
End Sub

Private Sub ApenasPCont_Click()
     Call objCT.ApenasPCont_Click
End Sub

Private Sub EntrePCont_Click()
     Call objCT.EntrePCont_Click
End Sub

Private Sub FaixaDataPCont_Click()
     Call objCT.FaixaDataPCont_Click
End Sub

Private Sub ApenasNComp_Click()
     Call objCT.ApenasNComp_Click
End Sub

Private Sub EntreNComp_Click()
     Call objCT.EntreNComp_Click
End Sub

Private Sub FaixaDataNComp_Click()
     Call objCT.FaixaDataNComp_Click
End Sub

Private Sub DataPContDe_Validate(Cancel As Boolean)
     Call objCT.DataPContDe_Validate(Cancel)
End Sub

Private Sub DataPContAte_Validate(Cancel As Boolean)
     Call objCT.DataPContAte_Validate(Cancel)
End Sub

Private Sub DataNCompDe_Validate(Cancel As Boolean)
     Call objCT.DataNCompDe_Validate(Cancel)
End Sub

Private Sub DataNCompAte_Validate(Cancel As Boolean)
     Call objCT.DataNCompAte_Validate(Cancel)
End Sub

Private Sub UpDownDataPContAte_DownClick()
     Call objCT.UpDownDataPContAte_DownClick
End Sub

Private Sub UpDownDataPContAte_UpClick()
     Call objCT.UpDownDataPContAte_UpClick
End Sub

Private Sub UpDownDataPContDe_DownClick()
     Call objCT.UpDownDataPContDe_DownClick
End Sub

Private Sub UpDownDataPContDe_UpClick()
     Call objCT.UpDownDataPContDe_UpClick
End Sub

Private Sub UpDownDataNCompAte_DownClick()
     Call objCT.UpDownDataNCompAte_DownClick
End Sub

Private Sub UpDownDataNCompAte_UpClick()
     Call objCT.UpDownDataNCompAte_UpClick
End Sub

Private Sub UpDownDataNCompDe_DownClick()
     Call objCT.UpDownDataNCompDe_DownClick
End Sub

Private Sub UpDownDataNCompDe_UpClick()
     Call objCT.UpDownDataNCompDe_UpClick
End Sub

Private Sub TiposCliente_Click()
     Call objCT.TiposCliente_Click
End Sub

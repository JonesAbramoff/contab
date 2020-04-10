VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ContatoClientePorEmailOcx 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   ScaleHeight     =   6900
   ScaleWidth      =   10005
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   6210
      Index           =   1
      Left            =   105
      TabIndex        =   105
      Top             =   615
      Width           =   9840
      Begin VB.ComboBox Modelo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ContatoClientePorEmail.ctx":0000
         Left            =   3225
         List            =   "ContatoClientePorEmail.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   15
         Width           =   5730
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   5925
         Left            =   0
         TabIndex        =   107
         Top             =   270
         Width           =   9810
         Begin VB.CheckBox IgnoraJaEnviados 
            Caption         =   "Ignorar clientes cujo email desse modelo já foi enviado"
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
            Left            =   6915
            TabIndex        =   6
            Top             =   150
            Width           =   2850
         End
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
            Height          =   390
            Left            =   5220
            TabIndex        =   5
            Top             =   150
            Width           =   1680
         End
         Begin VB.Frame Frame1 
            Caption         =   "Data Último Contato"
            Height          =   1620
            Index           =   0
            Left            =   4935
            TabIndex        =   160
            Top             =   1320
            Width           =   4770
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
               TabIndex        =   32
               Top             =   930
               Width           =   780
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
               TabIndex        =   29
               Top             =   600
               Width           =   1050
            End
            Begin VB.Frame FrameD1 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   330
               Index           =   0
               Left            =   840
               TabIndex        =   167
               Top             =   210
               Width           =   3765
               Begin MSMask.MaskEdBox DataContAte 
                  Height          =   300
                  Left            =   2295
                  TabIndex        =   27
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
                  TabIndex        =   26
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
                  TabIndex        =   25
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
                  TabIndex        =   28
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
                  Index           =   12
                  Left            =   1890
                  TabIndex        =   169
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
                  Index           =   1
                  Left            =   90
                  TabIndex        =   168
                  Top             =   75
                  Width           =   315
               End
            End
            Begin VB.Frame FrameD1 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   375
               Index           =   1
               Left            =   1050
               TabIndex        =   165
               Top             =   510
               Width           =   3390
               Begin VB.ComboBox ApenasQualifCont 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":003F
                  Left            =   240
                  List            =   "ContatoClientePorEmail.ctx":0049
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   45
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox ApenasDiasCont 
                  Height          =   315
                  Left            =   2085
                  TabIndex        =   31
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
                  TabIndex        =   166
                  Top             =   90
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD1 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   675
               Index           =   2
               Left            =   915
               TabIndex        =   161
               Top             =   900
               Width           =   3825
               Begin VB.ComboBox EntreQualifContDe 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":0064
                  Left            =   2220
                  List            =   "ContatoClientePorEmail.ctx":0071
                  Style           =   2  'Dropdown List
                  TabIndex        =   34
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.ComboBox EntreQualifContAte 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":008F
                  Left            =   2220
                  List            =   "ContatoClientePorEmail.ctx":009C
                  Style           =   2  'Dropdown List
                  TabIndex        =   36
                  Top             =   360
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox EntreDiasContDe 
                  Height          =   315
                  Left            =   375
                  TabIndex        =   33
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
                  TabIndex        =   35
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
                  Index           =   4
                  Left            =   1050
                  TabIndex        =   164
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
                  Index           =   6
                  Left            =   1050
                  TabIndex        =   163
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
                  Index           =   5
                  Left            =   3615
                  TabIndex        =   162
                  Top             =   45
                  Width           =   120
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
               TabIndex        =   24
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Data Próximo Contato"
            Height          =   1620
            Index           =   4
            Left            =   105
            TabIndex        =   150
            Top             =   1365
            Width           =   4770
            Begin VB.Frame FrameD4 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   675
               Index           =   2
               Left            =   915
               TabIndex        =   156
               Top             =   900
               Width           =   3825
               Begin VB.ComboBox EntreQualifPContAte 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":00BA
                  Left            =   2220
                  List            =   "ContatoClientePorEmail.ctx":00C7
                  Style           =   2  'Dropdown List
                  TabIndex        =   23
                  Top             =   360
                  Width           =   1365
               End
               Begin VB.ComboBox EntreQualifPContDe 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":00E5
                  Left            =   2220
                  List            =   "ContatoClientePorEmail.ctx":00F2
                  Style           =   2  'Dropdown List
                  TabIndex        =   21
                  Top             =   0
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox EntreDiasPContDe 
                  Height          =   315
                  Left            =   375
                  TabIndex        =   20
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
                  TabIndex        =   22
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
                  Index           =   20
                  Left            =   3615
                  TabIndex        =   159
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
                  Index           =   21
                  Left            =   1050
                  TabIndex        =   158
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
                  Index           =   22
                  Left            =   1050
                  TabIndex        =   157
                  Top             =   45
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD4 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   1
               Left            =   1020
               TabIndex        =   154
               Top             =   510
               Width           =   3540
               Begin VB.ComboBox ApenasQualifPCont 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":0110
                  Left            =   270
                  List            =   "ContatoClientePorEmail.ctx":011A
                  Style           =   2  'Dropdown List
                  TabIndex        =   17
                  Top             =   45
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox ApenasDiasPCont 
                  Height          =   315
                  Left            =   2115
                  TabIndex        =   18
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
                  Left            =   2820
                  TabIndex        =   155
                  Top             =   90
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD4 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   330
               Index           =   0
               Left            =   840
               TabIndex        =   151
               Top             =   210
               Width           =   3765
               Begin MSMask.MaskEdBox DataPContAte 
                  Height          =   300
                  Left            =   2295
                  TabIndex        =   14
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
                  TabIndex        =   13
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
                  TabIndex        =   12
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
                  TabIndex        =   15
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
                  Index           =   24
                  Left            =   90
                  TabIndex        =   153
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
                  Index           =   25
                  Left            =   1890
                  TabIndex        =   152
                  Top             =   75
                  Width           =   360
               End
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
               TabIndex        =   16
               Top             =   600
               Width           =   1050
            End
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
               TabIndex        =   19
               Top             =   930
               Width           =   780
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
               TabIndex        =   11
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Data Primeira Compra"
            Height          =   1635
            Index           =   2
            Left            =   105
            TabIndex        =   140
            Top             =   2970
            Width           =   4770
            Begin VB.Frame FrameD2 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   705
               Index           =   2
               Left            =   915
               TabIndex        =   146
               Top             =   900
               Width           =   3825
               Begin VB.ComboBox EntreQualifPCompAte 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":0135
                  Left            =   2220
                  List            =   "ContatoClientePorEmail.ctx":0142
                  Style           =   2  'Dropdown List
                  TabIndex        =   49
                  Top             =   360
                  Width           =   1365
               End
               Begin VB.ComboBox EntreQualifPCompDe 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":0160
                  Left            =   2220
                  List            =   "ContatoClientePorEmail.ctx":016D
                  Style           =   2  'Dropdown List
                  TabIndex        =   47
                  Top             =   0
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox EntreDiasPCompDe 
                  Height          =   315
                  Left            =   375
                  TabIndex        =   46
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
                  TabIndex        =   48
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
                  Index           =   2
                  Left            =   3615
                  TabIndex        =   149
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
                  Index           =   3
                  Left            =   1050
                  TabIndex        =   148
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
                  Index           =   7
                  Left            =   1050
                  TabIndex        =   147
                  Top             =   45
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD2 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   1
               Left            =   1005
               TabIndex        =   144
               Top             =   510
               Width           =   3630
               Begin VB.ComboBox ApenasQualifPComp 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":018B
                  Left            =   285
                  List            =   "ContatoClientePorEmail.ctx":0195
                  Style           =   2  'Dropdown List
                  TabIndex        =   43
                  Top             =   45
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox ApenasDiasPComp 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   44
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
                  Left            =   2835
                  TabIndex        =   145
                  Top             =   90
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD2 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   330
               Index           =   0
               Left            =   840
               TabIndex        =   141
               Top             =   210
               Width           =   3765
               Begin MSMask.MaskEdBox DataPCompAte 
                  Height          =   300
                  Left            =   2295
                  TabIndex        =   40
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
                  TabIndex        =   39
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
                  TabIndex        =   38
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
                  TabIndex        =   41
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
                  Index           =   9
                  Left            =   90
                  TabIndex        =   143
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
                  Index           =   11
                  Left            =   1890
                  TabIndex        =   142
                  Top             =   75
                  Width           =   360
               End
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
               TabIndex        =   42
               Top             =   600
               Width           =   1050
            End
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
               TabIndex        =   45
               Top             =   930
               Width           =   780
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
               TabIndex        =   37
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Data Última Compra"
            Height          =   1635
            Index           =   3
            Left            =   4920
            TabIndex        =   130
            Top             =   2970
            Width           =   4770
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
               TabIndex        =   58
               Top             =   930
               Width           =   780
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
               TabIndex        =   55
               Top             =   600
               Width           =   1155
            End
            Begin VB.Frame FrameD3 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   330
               Index           =   0
               Left            =   840
               TabIndex        =   137
               Top             =   210
               Width           =   3765
               Begin MSMask.MaskEdBox DataUCompAte 
                  Height          =   300
                  Left            =   2295
                  TabIndex        =   53
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
                  TabIndex        =   52
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
                  TabIndex        =   51
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
                  TabIndex        =   54
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
                  Index           =   13
                  Left            =   1890
                  TabIndex        =   139
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
                  Index           =   14
                  Left            =   90
                  TabIndex        =   138
                  Top             =   75
                  Width           =   315
               End
            End
            Begin VB.Frame FrameD3 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   705
               Index           =   2
               Left            =   915
               TabIndex        =   133
               Top             =   900
               Width           =   3825
               Begin VB.ComboBox EntreQualifUCompDe 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":01B0
                  Left            =   2220
                  List            =   "ContatoClientePorEmail.ctx":01BD
                  Style           =   2  'Dropdown List
                  TabIndex        =   60
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.ComboBox EntreQualifUCompAte 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":01DB
                  Left            =   2220
                  List            =   "ContatoClientePorEmail.ctx":01E8
                  Style           =   2  'Dropdown List
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox EntreDiasUCompDe 
                  Height          =   315
                  Left            =   375
                  TabIndex        =   59
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
                  TabIndex        =   61
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
                  Index           =   16
                  Left            =   1050
                  TabIndex        =   136
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
                  Index           =   17
                  Left            =   1050
                  TabIndex        =   135
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
                  Index           =   18
                  Left            =   3615
                  TabIndex        =   134
                  Top             =   45
                  Width           =   120
               End
            End
            Begin VB.Frame FrameD3 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   1
               Left            =   1095
               TabIndex        =   131
               Top             =   510
               Width           =   3390
               Begin VB.ComboBox ApenasQualifUComp 
                  Height          =   315
                  ItemData        =   "ContatoClientePorEmail.ctx":0206
                  Left            =   195
                  List            =   "ContatoClientePorEmail.ctx":0210
                  Style           =   2  'Dropdown List
                  TabIndex        =   56
                  Top             =   45
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox ApenasDiasuComp 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   57
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
                  TabIndex        =   132
                  Top             =   90
                  Width           =   480
               End
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
               TabIndex        =   50
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Não Efetuou Compra"
            Height          =   1290
            Index           =   5
            Left            =   105
            TabIndex        =   121
            Top             =   4590
            Width           =   4770
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
               TabIndex        =   68
               Top             =   615
               Width           =   975
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
               TabIndex        =   70
               Top             =   960
               Width           =   780
            End
            Begin VB.Frame FrameD5 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   330
               Index           =   2
               Left            =   1275
               TabIndex        =   127
               Top             =   915
               Width           =   3435
               Begin MSMask.MaskEdBox EntreDiasNCompDe 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   71
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
                  TabIndex        =   72
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
                  TabIndex        =   129
                  Top             =   60
                  Width           =   945
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
                  TabIndex        =   128
                  Top             =   60
                  Width           =   120
               End
            End
            Begin VB.Frame FrameD5 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   330
               Index           =   1
               Left            =   1260
               TabIndex        =   125
               Top             =   570
               Width           =   2070
               Begin MSMask.MaskEdBox ApenasDiasNComp 
                  Height          =   315
                  Left            =   15
                  TabIndex        =   69
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
                  TabIndex        =   126
                  Top             =   45
                  Width           =   945
               End
            End
            Begin VB.Frame FrameD5 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   300
               Index           =   0
               Left            =   900
               TabIndex        =   122
               Top             =   225
               Width           =   3630
               Begin MSMask.MaskEdBox DataNCompAte 
                  Height          =   300
                  Left            =   2235
                  TabIndex        =   66
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
                  TabIndex        =   65
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
                  TabIndex        =   64
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
                  TabIndex        =   67
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
                  TabIndex        =   124
                  Top             =   60
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
                  Index           =   31
                  Left            =   0
                  TabIndex        =   123
                  Top             =   60
                  Width           =   315
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
               TabIndex        =   63
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Tipos de Clientes"
            Height          =   1290
            Left            =   4920
            TabIndex        =   120
            Top             =   4590
            Width           =   4770
            Begin VB.ListBox TiposCliente 
               Height          =   960
               Left            =   135
               Style           =   1  'Checkbox
               TabIndex        =   73
               Top             =   255
               Width           =   4485
            End
         End
         Begin VB.CheckBox EmailValido 
            Caption         =   "Só trazer clientes com email válido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3210
            TabIndex        =   4
            Top             =   120
            Width           =   2025
         End
         Begin VB.ComboBox RespCallCenter 
            Height          =   315
            Left            =   1380
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   180
            Width           =   1695
         End
         Begin VB.Frame FrameCategoriaCliente 
            Caption         =   "Categoria de Cliente"
            Height          =   885
            Left            =   120
            TabIndex        =   108
            Top             =   495
            Width           =   9585
            Begin VB.ComboBox CategoriaCliente 
               Height          =   315
               Left            =   3105
               TabIndex        =   8
               Top             =   180
               Width           =   5745
            End
            Begin VB.ComboBox CategoriaClienteDe 
               Height          =   315
               Left            =   1290
               Sorted          =   -1  'True
               TabIndex        =   9
               Top             =   525
               Width           =   2760
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
               Left            =   90
               TabIndex        =   7
               Top             =   225
               Width           =   855
            End
            Begin VB.ComboBox CategoriaClienteAte 
               Height          =   315
               Left            =   6105
               Sorted          =   -1  'True
               TabIndex        =   10
               Top             =   510
               Width           =   2760
            End
            Begin VB.Label Label1 
               Caption         =   "Label5"
               Height          =   15
               Index           =   0
               Left            =   360
               TabIndex        =   112
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
               TabIndex        =   111
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
               Left            =   930
               TabIndex        =   110
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
               Left            =   2190
               TabIndex        =   109
               Top             =   225
               Width           =   855
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "R. Call Center:"
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
            Index           =   19
            Left            =   60
            TabIndex        =   113
            Top             =   210
            Width           =   1260
         End
      End
      Begin VB.Label LabelModelo 
         Caption         =   "Modelo de email:"
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
         Left            =   1680
         TabIndex        =   106
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6180
      Index           =   2
      Left            =   135
      TabIndex        =   89
      Top             =   615
      Visible         =   0   'False
      Width           =   9810
      Begin MSMask.MaskEdBox AnexoGrid 
         Height          =   255
         Left            =   165
         TabIndex        =   104
         Top             =   750
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Frame Frame7 
         Caption         =   "Email"
         Height          =   2745
         Left            =   1905
         TabIndex        =   99
         Top             =   3360
         Width           =   7785
         Begin VB.TextBox Anexo 
            Height          =   285
            Left            =   1260
            MaxLength       =   250
            TabIndex        =   85
            Top             =   1170
            Width           =   6465
         End
         Begin VB.TextBox Email 
            Height          =   285
            Left            =   1260
            MaxLength       =   8000
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   165
            Width           =   6465
         End
         Begin VB.TextBox Cc 
            Height          =   285
            Left            =   1260
            MaxLength       =   8000
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   495
            Width           =   6465
         End
         Begin VB.TextBox Assunto 
            Height          =   285
            Left            =   1260
            MaxLength       =   250
            TabIndex        =   84
            Top             =   825
            Width           =   6465
         End
         Begin VB.TextBox Mensagem 
            Enabled         =   0   'False
            Height          =   1155
            Left            =   1260
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   86
            Top             =   1515
            Width           =   6465
         End
         Begin VB.Label Label2 
            Caption         =   "Anexo:"
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
            Left            =   555
            TabIndex        =   103
            Top             =   1185
            Width           =   765
         End
         Begin VB.Label Label4 
            Caption         =   "Para:"
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
            Left            =   705
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   98
            Top             =   195
            Width           =   480
         End
         Begin VB.Label LabelCc 
            Caption         =   "Cc:"
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
            Left            =   855
            TabIndex        =   102
            Top             =   525
            Width           =   330
         End
         Begin VB.Label Label2 
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
            Index           =   1
            Left            =   420
            TabIndex        =   101
            Top             =   855
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Mensagem:"
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
            Left            =   210
            TabIndex        =   100
            Top             =   1500
            Width           =   1020
         End
      End
      Begin MSMask.MaskEdBox MensagemGrid 
         Height          =   255
         Left            =   1920
         TabIndex        =   97
         Top             =   2220
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AssuntoGrid 
         Height          =   255
         Left            =   3660
         TabIndex        =   96
         Top             =   2535
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CCGrid 
         Height          =   255
         Left            =   345
         TabIndex        =   95
         Top             =   1095
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox EmailGrid 
         Height          =   255
         Left            =   2490
         TabIndex        =   94
         Top             =   1950
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.ComboBox Carta 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ContatoClientePorEmail.ctx":022B
         Left            =   2340
         List            =   "ContatoClientePorEmail.ctx":0235
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   810
         Width           =   2235
      End
      Begin MSMask.MaskEdBox Filial 
         Height          =   255
         Left            =   3510
         TabIndex        =   92
         Top             =   1410
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   255
         Left            =   300
         TabIndex        =   91
         Top             =   1395
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CheckBox Selecionado 
         DragMode        =   1  'Automatic
         Height          =   270
         Left            =   1275
         TabIndex        =   90
         Top             =   1710
         Width           =   555
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   60
         Picture         =   "ContatoClientePorEmail.ctx":026A
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   3435
         Width           =   1725
      End
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   45
         Picture         =   "ContatoClientePorEmail.ctx":1284
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   4200
         Width           =   1725
      End
      Begin MSFlexGridLib.MSFlexGrid GridClientes 
         Height          =   630
         Left            =   15
         TabIndex        =   79
         Top             =   30
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   1111
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin MSMask.MaskEdBox ValorTotalCompras 
         Height          =   255
         Left            =   5115
         TabIndex        =   116
         Top             =   1185
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataUContato 
         Height          =   255
         Left            =   5130
         TabIndex        =   117
         Top             =   1560
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataPCompra 
         Height          =   255
         Left            =   6405
         TabIndex        =   118
         Top             =   1215
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataUCompra 
         Height          =   255
         Left            =   6450
         TabIndex        =   119
         Top             =   1560
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   525
      Left            =   2505
      TabIndex        =   114
      Top             =   0
      Width           =   4740
      Begin VB.ComboBox OpcoesTela 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   135
         Width           =   2760
      End
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
         Left            =   3810
         TabIndex        =   1
         Top             =   195
         Width           =   930
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
         Left            =   120
         TabIndex        =   115
         Top             =   195
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7320
      ScaleHeight     =   495
      ScaleWidth      =   2595
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "ContatoClientePorEmail.ctx":2466
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ContatoClientePorEmail.ctx":25F0
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Gravar Opção"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoEmail 
         Height          =   360
         Left            =   1110
         Picture         =   "ContatoClientePorEmail.ctx":274A
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Enviar email"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1635
         Picture         =   "ContatoClientePorEmail.ctx":30EC
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2145
         Picture         =   "ContatoClientePorEmail.ctx":361E
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6600
      Left            =   45
      TabIndex        =   87
      Top             =   255
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   11642
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clientes"
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
Attribute VB_Name = "ContatoClientePorEmailOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTContatoCliPorEmail
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoEmail_Click()
    Call objCT.BotaoEmail_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTContatoCliPorEmail
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

Private Sub DataNCompDe_Validate(Cancel As Boolean)
     Call objCT.DataNCompDe_Validate(Cancel)
End Sub

Private Sub DataNCompAte_Validate(Cancel As Boolean)
     Call objCT.DataNCompAte_Validate(Cancel)
End Sub

Private Sub DataPContDe_Validate(Cancel As Boolean)
     Call objCT.DataPContDe_Validate(Cancel)
End Sub

Private Sub DataPContAte_Validate(Cancel As Boolean)
     Call objCT.DataPContAte_Validate(Cancel)
End Sub

Private Sub TiposCliente_Click()
     Call objCT.TiposCliente_Click
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

Private Sub IgnoraJaEnviados_Click()
     Call objCT.IgnoraJaEnviados_Click
End Sub

Private Sub Modelo_Click()
     Call objCT.IgnoraJaEnviados_Click
End Sub

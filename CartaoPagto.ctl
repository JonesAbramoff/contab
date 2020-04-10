VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CartaoPagto 
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ScaleHeight     =   3810
   ScaleWidth      =   4770
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   1170
      Picture         =   "CartaoPagto.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3060
      Width           =   1035
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   555
      Left            =   2475
      Picture         =   "CartaoPagto.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3060
      Width           =   1035
   End
   Begin VB.TextBox Aprovacao 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   2535
      Width           =   2760
   End
   Begin VB.TextBox NumeroCartao 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1334
      Width           =   2760
   End
   Begin VB.ComboBox Parcelamento 
      Height          =   315
      ItemData        =   "CartaoPagto.ctx":025C
      Left            =   1560
      List            =   "CartaoPagto.ctx":025E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   727
      Width           =   2760
   End
   Begin VB.ComboBox Adm 
      Height          =   315
      ItemData        =   "CartaoPagto.ctx":0260
      Left            =   1560
      List            =   "CartaoPagto.ctx":0262
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2760
   End
   Begin MSMask.MaskEdBox Validade 
      Height          =   300
      Left            =   1560
      TabIndex        =   6
      Top             =   1941
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   5
      Format          =   "mm/yyyy"
      Mask            =   "##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Aprovação:"
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
      Left            =   495
      TabIndex        =   9
      Top             =   2595
      Width           =   990
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Válido Até:"
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
      TabIndex        =   7
      Top             =   1965
      Width           =   945
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   765
      TabIndex        =   5
      Top             =   1395
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Parcelamento:"
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
      Left            =   255
      TabIndex        =   3
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cartão:"
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
      Left            =   855
      TabIndex        =   2
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "CartaoPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Validade_Change()

End Sub
